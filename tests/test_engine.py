"""test_engine.py — Regression tests for the detached Python borrowing base engine.

Test structure:
  TestEligibilityUnit   — 10 eligibility tests, synthetic data, no workbook
  TestCollateralTierUnit — collateral tier lookup, synthetic data, no workbook
  TestAdvanceRateUnit    — advance rate bucketing, synthetic data, no workbook
  TestVaeHandlingUnit    — VAE override logic, synthetic WorkbookData
  TestScenarioUnit       — scenario record construction + injection
  TestIntegration        — end-to-end against real workbook (skipped if absent)

Usage:
    cd C:\\Users\\henry.yan\\borrowing-base
    PYTHONPATH=src python -m pytest tests/test_engine.py -v
    # or
    PYTHONPATH=src python tests/test_engine.py
"""

from __future__ import annotations

import sys
import unittest
from datetime import date
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "src"))

from borrowing_base_workbench.calculator import (
    EligibilityResult,
    _lookup_collateral_tier,
    _pick_advance_rate,
    calculate_asset,
    calculate_portfolio,
)
from borrowing_base_workbench.loader import (
    AvailabilityMeta,
    CollateralTier,
    LoanRecord,
    ManualPortfolioFlags,
    PolicyConfig,
    VaeRecord,
    WorkbookData,
)
from borrowing_base_workbench.scenario import build_scenario_records, inject_scenario

# ---------------------------------------------------------------------------
# Workbook path (Henry Yan's machine)
# ---------------------------------------------------------------------------

WORKBOOK = Path(r"C:\Users\henry.yan\Downloads\2025-02-11_BDC Borrowing_Base_v8.xlsm")

# Reference values verified against the 2025-02-11 workbook (measurement 2025-12-31)
REF = {
    "asset_count":                   41,
    "eligible_count":                25,
    "ineligible_count":              16,
    "vae_affected_count":            11,
    "total_portfolio_par":           330_231_191,
    "total_pre_conc_eligible_value": 120_210_997,
    "total_borrowing_value":         73_085_307,
    "raw_waar":                      0.6080,
    "discrete_obligor_count":        25,
    "waar_cap":                      0.65,
    "final_waar":                    0.6080,
    "total_excess":                  29_281_364,
    "net_abv":                       90_929_633,
    "test_a_facility":               200_000_000,
    "test_b_borrowing_base":         60_753_941,
    "test_c_credit_enhancement":     81_704_750,
    "availability":                  60_753_941,
    # Per-test concentration excesses
    "obligor_excess":                24_911_002,
    "industry_excess":               4_370_362,
}


# ---------------------------------------------------------------------------
# Synthetic data factories (no workbook required)
# ---------------------------------------------------------------------------

def _make_tiers() -> list[CollateralTier]:
    return [
        CollateralTier(max_first_lien_lev=3.25, max_filo_2l_lev=3.75, applicable_pct=1.00),
        CollateralTier(max_first_lien_lev=4.00, max_filo_2l_lev=4.25, applicable_pct=0.925),
        CollateralTier(max_first_lien_lev=5.00, max_filo_2l_lev=5.00, applicable_pct=0.85),
        CollateralTier(max_first_lien_lev=999.0, max_filo_2l_lev=999.0, applicable_pct=0.0),
    ]


def _make_policy(**kwargs) -> PolicyConfig:
    defaults = dict(
        max_total_leverage_at_inclusion=6.5,
        min_ebitda_second_lien_mm=10.0,
        min_acquisition_price=0.90,
        max_original_maturity_years=7.0,
        rate_first_lien_large=0.675,
        rate_first_lien_mid=0.675,
        rate_first_lien_small=0.60,
        rate_filo_low_lev=0.65,
        rate_filo_mid_lev=0.60,
        rate_filo_high_lev=0.55,
        rate_second_lien=0.55,
        cap_high_diversity=0.65,
        cap_mid_diversity=0.60,
        cap_low_diversity=0.50,
        threshold_high_diversity=20,
        threshold_mid_diversity=12,
        collateral_tiers=_make_tiers(),
        conc_max_2l_filo_high_lev=0.20,
        conc_max_second_lien=0.20,
        conc_max_non_first_lien=0.25,
        conc_max_small_ebitda=0.20,
        conc_max_obligor=0.075,
        conc_max_largest_industry=0.20,
        conc_max_second_industry=0.15,
        conc_max_other_industries=0.10,
        conc_fixed_rate=0.20,
        conc_limited_industry=0.10,
        conc_max_ddtl_revolver=0.15,
        conc_max_non_sponsor=0.15,
        conc_max_div_recap=0.10,
        limited_industries=[],
        industry_taxonomy=["Software", "Healthcare", "Retail", "Business Services",
                           "Manufacturing", "Technology"],
    )
    defaults.update(kwargs)
    return PolicyConfig(**defaults)


def _make_flags(name: str, **kwargs) -> ManualPortfolioFlags:
    return ManualPortfolioFlags(
        obligor_name=name,
        limited_industry=kwargs.get("limited_industry", False),
        non_sponsor=kwargs.get("non_sponsor", False),
        div_recap=kwargs.get("div_recap", False),
        is_dip=kwargs.get("is_dip", False),
        agent_addback_discretion=False,
        agent_post_inclusion_haircut=False,
        agent_addback_haircut_pct=None,
    )


def _make_loan(name: str = "Test Co", row: int = 1, **kwargs) -> LoanRecord:
    olb = kwargs.get("olb", 10_000_000.0)
    return LoanRecord(
        source_row=row,
        ccm_id="TEST",
        fund_id="TEST",
        loan_id=row,
        obligor_name=name,
        security_display="1st Lien",
        loan_type=kwargs.get("loan_type", "First Lien"),
        security_type="Term Loan",
        denomination=kwargs.get("denomination", "USD"),
        commitment_balance=olb,
        olb=olb,
        unfunded=0.0,
        purchase_price=kwargs.get("purchase_price", 1.0),
        fmv=kwargs.get("fmv", olb),
        country=kwargs.get("country", "United States"),
        investment_date=kwargs.get("investment_date", date(2023, 1, 1)),
        maturity_date=kwargs.get("maturity_date", date(2029, 1, 1)),  # 6-year term
        rate_type=kwargs.get("rate_type", "Floating"),
        industry=kwargs.get("industry", "Software"),
        payment_frequency="M",
        cov_lite=kwargs.get("cov_lite", False),
        pik_pct=kwargs.get("pik_pct", 0.0),
        spread=0.055,
        sofr_floor=0.005,
        is_ddtl=(kwargs.get("loan_type", "First Lien") == "DDTL"),
        inception_date=date(2023, 1, 1),
        inclusion_date=kwargs.get("inclusion_date", date(2023, 1, 1)),
        initial_ebitda_mm=25_000_000.0,
        initial_ebitda_nonadj_mm=25_000_000.0,
        initial_senior_leverage=kwargs.get("net_detachment", 2.0),
        initial_total_leverage=kwargs.get("initial_total_leverage", 2.0),
        initial_interest_coverage=2.5,
        current_ebitda_mm=kwargs.get("current_ebitda_mm", 25.0),  # millions
        current_senior_leverage=kwargs.get("net_detachment", 2.0),
        current_interest_coverage=2.5,
        net_detachment=kwargs.get("net_detachment", 2.0),
        net_attachment=kwargs.get("net_attachment", 2.5),
    )


def _make_data(loans: list[LoanRecord], vaes: list[VaeRecord] | None = None,
               flags: dict | None = None, policy: PolicyConfig | None = None) -> WorkbookData:
    return WorkbookData(
        workbook_path="synthetic",
        loans=loans,
        obligors=[],
        vaes=vaes or [],
        portfolio_flags=flags or {},
        policy=policy or _make_policy(),
        availability_meta=AvailabilityMeta(
            facility_amount=200_000_000.0,
            available_collections=0.0,
            current_advances=50_000_000.0,
            measurement_date=date(2025, 12, 31),
        ),
    )


# ---------------------------------------------------------------------------
# TestEligibilityUnit — 10 eligibility tests in isolation
# ---------------------------------------------------------------------------

class TestEligibilityUnit(unittest.TestCase):
    """Tests for EligibilityResult.compute() — each test exercises one fail path."""

    def _base_loan(self, **kwargs) -> LoanRecord:
        return _make_loan("Acme", **kwargs)

    def _base_flags(self, **kwargs) -> ManualPortfolioFlags:
        return _make_flags("Acme", **kwargs)

    def _compute(self, loan: LoanRecord, flags=None, vaes=None) -> EligibilityResult:
        return EligibilityResult.compute(
            loan,
            flags or self._base_flags(),
            _make_policy(),
            vaes or [],
        )

    # ---- Happy path --------------------------------------------------------

    def test_eligible_first_lien_happy_path(self):
        result = self._compute(self._base_loan())
        self.assertTrue(result.eligible)
        self.assertEqual(result.failed_tests, [])

    def test_eligible_filo_happy_path(self):
        result = self._compute(self._base_loan(loan_type="FILO", net_detachment=1.0, net_attachment=2.0))
        self.assertTrue(result.eligible)

    def test_eligible_second_lien_large_ebitda(self):
        result = self._compute(self._base_loan(loan_type="Second Lien", current_ebitda_mm=15.0))
        self.assertTrue(result.eligible)

    # ---- Test A: invalid loan type -----------------------------------------

    def test_fail_invalid_loan_type(self):
        result = self._compute(self._base_loan(loan_type="Mezzanine"))
        self.assertFalse(result.eligible)
        self.assertFalse(result.test_valid_type)
        self.assertTrue(any("type" in f.lower() for f in result.failed_tests))

    # ---- Test B: non-USD ---------------------------------------------------

    def test_fail_non_usd(self):
        result = self._compute(self._base_loan(denomination="CAD"))
        self.assertFalse(result.eligible)
        self.assertFalse(result.test_usd_denomination)

    def test_pass_usd_case_insensitive(self):
        result = self._compute(self._base_loan(denomination="usd"))
        self.assertTrue(result.test_usd_denomination)

    # ---- Test C: non-US domicile -------------------------------------------

    def test_fail_non_us(self):
        result = self._compute(self._base_loan(country="Canada"))
        self.assertFalse(result.eligible)
        self.assertFalse(result.test_us_domicile)

    def test_pass_us_variants(self):
        for country in ("United States", "US", "USA", "U.S.", "UNITED STATES"):
            with self.subTest(country=country):
                result = self._compute(self._base_loan(country=country))
                self.assertTrue(result.test_us_domicile)

    # ---- Test D: cov-lite --------------------------------------------------

    def test_fail_cov_lite(self):
        result = self._compute(self._base_loan(cov_lite=True))
        self.assertFalse(result.eligible)
        self.assertFalse(result.test_not_cov_lite)

    # ---- Test E: PIK and DIP -----------------------------------------------

    def test_fail_pik_pct(self):
        result = self._compute(self._base_loan(pik_pct=0.05))
        self.assertFalse(result.eligible)
        self.assertFalse(result.test_non_pik_dip)

    def test_fail_dip_flag(self):
        result = self._compute(self._base_loan(), flags=self._base_flags(is_dip=True))
        self.assertFalse(result.eligible)
        self.assertFalse(result.test_non_pik_dip)

    def test_pass_zero_pik(self):
        result = self._compute(self._base_loan(pik_pct=0.0))
        self.assertTrue(result.test_non_pik_dip)

    # ---- Test F: acquisition price -----------------------------------------

    def test_fail_low_purchase_price(self):
        result = self._compute(self._base_loan(purchase_price=0.85))  # < 0.90 min
        self.assertFalse(result.eligible)
        self.assertFalse(result.test_acquisition_price)

    def test_pass_at_min_price(self):
        result = self._compute(self._base_loan(purchase_price=0.90))
        self.assertTrue(result.test_acquisition_price)

    # ---- Test G: 2L EBITDA -------------------------------------------------

    def test_fail_second_lien_small_ebitda(self):
        result = self._compute(self._base_loan(loan_type="Second Lien", current_ebitda_mm=8.0))
        self.assertFalse(result.eligible)
        self.assertFalse(result.test_second_lien_ebitda)

    def test_pass_fl_ebitda_not_checked(self):
        """Test G is N/A (True) for First Lien regardless of EBITDA size."""
        result = self._compute(self._base_loan(loan_type="First Lien", current_ebitda_mm=2.0))
        self.assertTrue(result.test_second_lien_ebitda)

    # ---- Test H: total leverage --------------------------------------------

    def test_fail_high_initial_leverage(self):
        result = self._compute(self._base_loan(initial_total_leverage=7.0))
        self.assertFalse(result.eligible)
        self.assertFalse(result.test_total_leverage)

    def test_pass_null_leverage(self):
        """Test H: None leverage → pass (data absent for legacy loans)."""
        loan = _make_loan("Acme", initial_total_leverage=None)
        loan = LoanRecord(
            **{**loan.__dict__, "initial_total_leverage": None}
        )
        result = self._compute(loan)
        self.assertTrue(result.test_total_leverage)

    # ---- Test I: original maturity -----------------------------------------

    def test_fail_too_long_maturity(self):
        result = self._compute(self._base_loan(
            inclusion_date=date(2023, 1, 1),
            maturity_date=date(2031, 1, 1),  # 8 years > 7-year limit
        ))
        self.assertFalse(result.eligible)
        self.assertFalse(result.test_original_maturity)

    def test_pass_exactly_at_limit(self):
        # 2023-01-01 to 2029-12-29 = 2554 days / 365 = 6.997 years (<= 7.0 limit)
        result = self._compute(self._base_loan(
            inclusion_date=date(2023, 1, 1),
            maturity_date=date(2029, 12, 29),
        ))
        self.assertTrue(result.test_original_maturity)

    def test_fail_missing_maturity_date(self):
        loan = _make_loan("Acme", maturity_date=None)
        loan = LoanRecord(**{**loan.__dict__, "maturity_date": None})
        result = self._compute(loan)
        self.assertFalse(result.test_original_maturity)

    # ---- Test J: bankruptcy VAE --------------------------------------------

    def test_fail_bankruptcy_vae(self):
        vae = VaeRecord(
            source_row=1, borrower="Acme", event_type="Chapter 11 Bankruptcy",
            specific_test="", material_modification="", vae_date=date(2024, 1, 1),
            ebitda_at_vae=None, agent_adj_haircut=None, permitted_ebitda=None,
            senior_debt=None, total_debt=None, net_senior_leverage=None,
            net_total_leverage=None, interest_coverage_at_vae=None,
            vae_agent_assigned_value=None,
        )
        result = self._compute(self._base_loan(), vaes=[vae])
        self.assertFalse(result.eligible)
        self.assertFalse(result.test_not_defaulted)

    def test_pass_vae_without_bankruptcy(self):
        vae = VaeRecord(
            source_row=1, borrower="Acme", event_type="Revenue Decline",
            specific_test="", material_modification="", vae_date=date(2024, 1, 1),
            ebitda_at_vae=None, agent_adj_haircut=None, permitted_ebitda=None,
            senior_debt=None, total_debt=None, net_senior_leverage=None,
            net_total_leverage=None, interest_coverage_at_vae=None,
            vae_agent_assigned_value=1_000_000.0,
        )
        result = self._compute(self._base_loan(), vaes=[vae])
        self.assertTrue(result.test_not_defaulted)  # no bankruptcy keyword

    def test_multiple_failures_reported(self):
        """Multiple test failures should each appear in failed_tests list."""
        result = self._compute(self._base_loan(
            denomination="CAD",
            cov_lite=True,
            pik_pct=0.05,
        ))
        self.assertFalse(result.eligible)
        self.assertGreaterEqual(len(result.failed_tests), 3)


# ---------------------------------------------------------------------------
# TestCollateralTierUnit
# ---------------------------------------------------------------------------

class TestCollateralTierUnit(unittest.TestCase):
    """Tests for _lookup_collateral_tier() — leverage → tier percentage."""

    def _tiers(self) -> list[CollateralTier]:
        return _make_tiers()

    def test_fl_tier1_low_leverage(self):
        loan = _make_loan("A", loan_type="First Lien", net_detachment=2.0)
        tier, pct = _lookup_collateral_tier(loan, self._tiers())
        self.assertEqual(pct, 1.0)

    def test_fl_tier1_at_threshold(self):
        loan = _make_loan("A", loan_type="First Lien", net_detachment=3.25)
        _, pct = _lookup_collateral_tier(loan, self._tiers())
        self.assertEqual(pct, 1.0)

    def test_fl_tier2_just_above_tier1(self):
        loan = _make_loan("A", loan_type="First Lien", net_detachment=3.50)
        _, pct = _lookup_collateral_tier(loan, self._tiers())
        self.assertAlmostEqual(pct, 0.925)

    def test_fl_tier3_leverage(self):
        loan = _make_loan("A", loan_type="First Lien", net_detachment=4.50)
        _, pct = _lookup_collateral_tier(loan, self._tiers())
        self.assertAlmostEqual(pct, 0.85)

    def test_fl_tier4_above_5x(self):
        loan = _make_loan("A", loan_type="First Lien", net_detachment=5.50)
        _, pct = _lookup_collateral_tier(loan, self._tiers())
        self.assertEqual(pct, 0.0)

    def test_filo_uses_net_attachment(self):
        """FILO tiers compare net_attachment, not net_detachment.
        net_detachment=1.0 would hit Tier 1 if we used it (< 3.25 FL threshold).
        net_attachment=4.0 puts this in FILO Tier 2 (3.75 < 4.0 <= 4.25).
        """
        loan = _make_loan("A", loan_type="FILO", net_detachment=1.0, net_attachment=4.0)
        _, pct = _lookup_collateral_tier(loan, self._tiers())
        self.assertAlmostEqual(pct, 0.925)  # Tier 2 for FILO

    def test_filo_tier1_low_attachment(self):
        loan = _make_loan("A", loan_type="FILO", net_detachment=0.5, net_attachment=2.0)
        _, pct = _lookup_collateral_tier(loan, self._tiers())
        self.assertEqual(pct, 1.0)

    def test_null_leverage_returns_zero(self):
        loan = _make_loan("A", loan_type="First Lien", net_detachment=None)
        loan = LoanRecord(**{**loan.__dict__, "net_detachment": None})
        tier, pct = _lookup_collateral_tier(loan, self._tiers())
        self.assertIsNone(tier)
        self.assertEqual(pct, 0.0)

    def test_second_lien_uses_net_attachment(self):
        """Second Lien tiers use net_attachment (same as FILO).
        net_attachment=4.0 puts this in Tier 2 (3.75 < 4.0 <= 4.25).
        """
        loan = _make_loan("A", loan_type="Second Lien", net_detachment=1.0, net_attachment=4.0)
        _, pct = _lookup_collateral_tier(loan, self._tiers())
        self.assertAlmostEqual(pct, 0.925)


# ---------------------------------------------------------------------------
# TestAdvanceRateUnit
# ---------------------------------------------------------------------------

class TestAdvanceRateUnit(unittest.TestCase):
    """Tests for _pick_advance_rate()."""

    def _p(self) -> PolicyConfig:
        return _make_policy()

    def test_fl_large_ebitda(self):
        loan = _make_loan("A", loan_type="First Lien", current_ebitda_mm=25.0)
        label, rate = _pick_advance_rate(loan, self._p())
        self.assertIn("$20MM", label)
        self.assertAlmostEqual(rate, 0.675)

    def test_fl_mid_ebitda(self):
        loan = _make_loan("A", loan_type="First Lien", current_ebitda_mm=15.0)
        label, rate = _pick_advance_rate(loan, self._p())
        self.assertIn("$10", label)
        self.assertAlmostEqual(rate, 0.675)

    def test_fl_small_ebitda(self):
        loan = _make_loan("A", loan_type="First Lien", current_ebitda_mm=8.0)
        label, rate = _pick_advance_rate(loan, self._p())
        self.assertIn("$10MM", label)
        self.assertAlmostEqual(rate, 0.60)

    def test_filo_low_lev(self):
        loan = _make_loan("A", loan_type="FILO", net_detachment=0.5)
        label, rate = _pick_advance_rate(loan, self._p())
        self.assertAlmostEqual(rate, 0.65)

    def test_filo_mid_lev(self):
        loan = _make_loan("A", loan_type="FILO", net_detachment=1.0)
        label, rate = _pick_advance_rate(loan, self._p())
        self.assertAlmostEqual(rate, 0.60)

    def test_filo_high_lev(self):
        loan = _make_loan("A", loan_type="FILO", net_detachment=2.0)
        label, rate = _pick_advance_rate(loan, self._p())
        self.assertAlmostEqual(rate, 0.55)

    def test_filo_null_leverage_conservative(self):
        loan = _make_loan("A", loan_type="FILO", net_detachment=None)
        loan = LoanRecord(**{**loan.__dict__, "net_detachment": None})
        label, rate = _pick_advance_rate(loan, self._p())
        self.assertAlmostEqual(rate, 0.55)  # most conservative

    def test_second_lien_fixed_rate(self):
        loan = _make_loan("A", loan_type="Second Lien")
        label, rate = _pick_advance_rate(loan, self._p())
        self.assertAlmostEqual(rate, 0.55)

    def test_ddtl_uses_fl_ebitda_buckets(self):
        loan = _make_loan("A", loan_type="DDTL", current_ebitda_mm=25.0,
                          net_detachment=2.0, net_attachment=2.0)
        loan = LoanRecord(**{**loan.__dict__, "loan_type": "DDTL", "is_ddtl": True})
        label, rate = _pick_advance_rate(loan, self._p())
        self.assertAlmostEqual(rate, 0.675)


# ---------------------------------------------------------------------------
# TestVaeHandlingUnit
# ---------------------------------------------------------------------------

class TestVaeHandlingUnit(unittest.TestCase):
    """Tests for VAE assigned value override in calculate_asset()."""

    def _run(self, loan: LoanRecord, vaes: list[VaeRecord] | None = None,
             flags: ManualPortfolioFlags | None = None) -> object:
        data = _make_data(
            [loan],
            vaes=vaes or [],
            flags={loan.obligor_name: flags or _make_flags(loan.obligor_name)},
        )
        return calculate_portfolio(data).assets[0]

    def test_no_vae_uses_min_of_fmv_and_tier(self):
        """Without VAE: pre_conc_value = min(fmv, tier_value)."""
        loan = _make_loan("A", olb=10_000_000, fmv=8_000_000, net_detachment=2.0)
        result = self._run(loan)
        tier_val = 1.0 * 10_000_000  # Tier 1 (100%)
        self.assertAlmostEqual(result.pre_conc_value, min(8_000_000, tier_val))

    def test_vae_reduces_assigned_value(self):
        """VAE assigned value below FMV and tier value → VAE wins."""
        loan = _make_loan("A", olb=10_000_000, fmv=10_000_000, net_detachment=2.0)
        vae = VaeRecord(
            source_row=1, borrower="A", event_type="Revenue Decline",
            specific_test="", material_modification="", vae_date=date(2024, 6, 1),
            ebitda_at_vae=None, agent_adj_haircut=None, permitted_ebitda=None,
            senior_debt=None, total_debt=None, net_senior_leverage=None,
            net_total_leverage=None, interest_coverage_at_vae=None,
            vae_agent_assigned_value=3_000_000.0,  # well below fmv/tier
        )
        result = self._run(loan, vaes=[vae])
        self.assertTrue(result.has_vae)
        self.assertAlmostEqual(result.pre_conc_value, 3_000_000.0)

    def test_vae_does_not_increase_above_fmv(self):
        """VAE assigned value above FMV → still capped by min()."""
        loan = _make_loan("A", olb=10_000_000, fmv=6_000_000, net_detachment=2.0)
        vae = VaeRecord(
            source_row=1, borrower="A", event_type="Revenue Decline",
            specific_test="", material_modification="", vae_date=date(2024, 1, 1),
            ebitda_at_vae=None, agent_adj_haircut=None, permitted_ebitda=None,
            senior_debt=None, total_debt=None, net_senior_leverage=None,
            net_total_leverage=None, interest_coverage_at_vae=None,
            vae_agent_assigned_value=12_000_000.0,  # above FMV
        )
        result = self._run(loan, vaes=[vae])
        self.assertAlmostEqual(result.pre_conc_value, 6_000_000.0)  # FMV wins

    def test_most_recent_vae_used(self):
        """The VAE with the latest vae_date and non-None value is used."""
        loan = _make_loan("A", olb=10_000_000, fmv=10_000_000, net_detachment=2.0)
        old_vae = VaeRecord(
            source_row=1, borrower="A", event_type="Revenue Decline",
            specific_test="", material_modification="", vae_date=date(2023, 1, 1),
            ebitda_at_vae=None, agent_adj_haircut=None, permitted_ebitda=None,
            senior_debt=None, total_debt=None, net_senior_leverage=None,
            net_total_leverage=None, interest_coverage_at_vae=None,
            vae_agent_assigned_value=8_000_000.0,  # older, higher value
        )
        new_vae = VaeRecord(
            source_row=2, borrower="A", event_type="Revenue Decline",
            specific_test="", material_modification="", vae_date=date(2024, 6, 1),
            ebitda_at_vae=None, agent_adj_haircut=None, permitted_ebitda=None,
            senior_debt=None, total_debt=None, net_senior_leverage=None,
            net_total_leverage=None, interest_coverage_at_vae=None,
            vae_agent_assigned_value=2_000_000.0,  # newer, lower value
        )
        result = self._run(loan, vaes=[old_vae, new_vae])
        self.assertAlmostEqual(result.pre_conc_value, 2_000_000.0)

    def test_ineligible_loan_zero_pre_conc(self):
        """Ineligible loans (PIK) get pre_conc_value = 0."""
        loan = _make_loan("A", pik_pct=0.05)  # fails test E
        result = self._run(loan)
        self.assertFalse(result.eligibility.eligible)
        self.assertEqual(result.pre_conc_value, 0.0)
        self.assertEqual(result.borrowing_value, 0.0)


# ---------------------------------------------------------------------------
# TestScenarioUnit — scenario record construction and injection
# ---------------------------------------------------------------------------

class TestScenarioUnit(unittest.TestCase):
    """Tests for build_scenario_records() and inject_scenario()."""

    _BASE_SCENARIO = {
        "company_name":            "Test Holdco",
        "security_type":           "First Lien",
        "ltm_revenue":             "100000000",
        "ltm_adj_ebitda":          "25000000",
        "drawn_revolver":          "0",
        "first_out_balance":       "0",
        "pari_passu":              "0",
        "bdc_balance":             "20000000",
        "total_sm_balance":        "20000000",
        "cash_balance":            "0",
        "interest_coverage":       "2.5",
        "loan_denomination":       "USD",
        "purchase_price":          "1.0",
        "country":                 "United States",
        "investment_date":         "2025-01-15",
        "maturity_date":           "2031-01-15",
        "rate_type":               "Floating",
        "industry_classification": "Software",
        "payment_frequency":       "M",
        "pik_pct":                 "0",
        "spread":                  "0.055",
        "sofr_floor":              "0.005",
        "attach_point":            "0.0",
    }

    def _scenario(self, **overrides) -> dict:
        s = dict(self._BASE_SCENARIO)
        s.update(overrides)
        return s

    def test_build_first_lien_defaults(self):
        loan, obligor, flags = build_scenario_records(self._scenario())
        self.assertEqual(loan.obligor_name, "Test Holdco")
        self.assertEqual(loan.loan_type, "First Lien")
        self.assertEqual(loan.denomination, "USD")
        self.assertEqual(loan.country, "United States")
        self.assertFalse(loan.cov_lite)
        self.assertFalse(loan.pik_pct > 0)

    def test_build_obligor_balances(self):
        loan, obligor, flags = build_scenario_records(self._scenario())
        self.assertAlmostEqual(obligor.debt_balance, 20_000_000)
        self.assertAlmostEqual(obligor.ltm_adj_ebitda, 25_000_000)
        self.assertAlmostEqual(obligor.ltm_revenue, 100_000_000)

    def test_build_fmv_at_par(self):
        loan, obligor, flags = build_scenario_records(self._scenario(purchase_price="1.0"))
        self.assertAlmostEqual(loan.fmv, 20_000_000)  # par × OLB

    def test_build_fmv_discounted(self):
        loan, obligor, flags = build_scenario_records(self._scenario(purchase_price="0.95"))
        self.assertAlmostEqual(loan.fmv, 19_000_000)  # 0.95 × 20M

    def test_build_leverage_derived_from_financials(self):
        """net_detachment = (revolver + first_out - cash) / ebitda."""
        scenario = self._scenario(
            drawn_revolver="10000000", first_out_balance="5000000",
            cash_balance="2000000", ltm_adj_ebitda="25000000",
        )
        loan, obligor, flags = build_scenario_records(scenario)
        expected_det = (10_000_000 + 5_000_000 - 2_000_000) / 25_000_000  # 0.52x
        self.assertAlmostEqual(loan.net_detachment, expected_det)
        self.assertAlmostEqual(obligor.net_detachment, expected_det)

    def test_build_net_attachment_formula(self):
        """net_attachment includes pari_passu and total_sm_balance."""
        scenario = self._scenario(
            drawn_revolver="0", first_out_balance="0", pari_passu="5000000",
            total_sm_balance="20000000", cash_balance="0",
            ltm_adj_ebitda="25000000",
        )
        loan, obligor, flags = build_scenario_records(scenario)
        expected_att = (0 + 0 + 5_000_000 + 20_000_000 - 0) / 25_000_000  # 1.0x
        self.assertAlmostEqual(loan.net_attachment, expected_att)

    def test_build_current_ebitda_in_millions(self):
        """LoanRecord.current_ebitda_mm is in millions (not dollars)."""
        loan, _, _ = build_scenario_records(self._scenario(ltm_adj_ebitda="25000000"))
        self.assertAlmostEqual(loan.current_ebitda_mm, 25.0)  # 25.0 MM

    def test_build_inclusion_date_equals_investment_date(self):
        loan, _, _ = build_scenario_records(self._scenario(investment_date="2025-06-01"))
        self.assertEqual(str(loan.investment_date), "2025-06-01")
        self.assertEqual(loan.inclusion_date, loan.investment_date)

    def test_build_flags_all_false(self):
        _, _, flags = build_scenario_records(self._scenario())
        self.assertFalse(flags.limited_industry)
        self.assertFalse(flags.non_sponsor)
        self.assertFalse(flags.div_recap)
        self.assertFalse(flags.is_dip)

    def test_build_filo_type(self):
        loan, _, _ = build_scenario_records(self._scenario(security_type="FILO"))
        self.assertEqual(loan.loan_type, "FILO")

    def test_build_attach_point_fallback_when_ebitda_zero(self):
        """attach_point used as net_detachment fallback when EBITDA = 0."""
        scenario = self._scenario(ltm_adj_ebitda="0", attach_point="3.5")
        loan, obligor, _ = build_scenario_records(scenario)
        self.assertAlmostEqual(loan.net_detachment, 3.5)
        self.assertAlmostEqual(obligor.net_detachment, 3.5)

    def test_inject_adds_one_loan(self):
        data = _make_data([_make_loan("Existing")])
        loan, obligor, flags = build_scenario_records(self._scenario())
        augmented = inject_scenario(data, loan, obligor, flags)
        self.assertEqual(len(augmented.loans), 2)
        self.assertEqual(augmented.loans[-1].obligor_name, "Test Holdco")

    def test_inject_does_not_mutate_original(self):
        data = _make_data([_make_loan("Existing")])
        loan, obligor, flags = build_scenario_records(self._scenario())
        augmented = inject_scenario(data, loan, obligor, flags)
        # Original WorkbookData untouched
        self.assertEqual(len(data.loans), 1)
        self.assertNotIn("Test Holdco", data.portfolio_flags)

    def test_inject_flags_registered(self):
        data = _make_data([_make_loan("Existing")])
        loan, obligor, flags = build_scenario_records(self._scenario())
        augmented = inject_scenario(data, loan, obligor, flags)
        self.assertIn("Test Holdco", augmented.portfolio_flags)

    def test_inject_obligor_registered(self):
        data = _make_data([_make_loan("Existing")])
        loan, obligor, flags = build_scenario_records(self._scenario())
        augmented = inject_scenario(data, loan, obligor, flags)
        self.assertEqual(len(augmented.obligors), 1)  # was 0, now 1

    def test_scenario_increases_portfolio_count(self):
        """After injection, calculate_portfolio sees one more asset."""
        data = _make_data([_make_loan("Existing")])
        loan, obligor, flags = build_scenario_records(self._scenario())
        augmented = inject_scenario(data, loan, obligor, flags)
        result = calculate_portfolio(augmented)
        self.assertEqual(len(result.assets), 2)

    def test_scenario_loan_eligible_by_default(self):
        """A clean scenario loan with valid inputs should be eligible."""
        data = _make_data([])
        loan, obligor, flags = build_scenario_records(self._scenario())
        augmented = inject_scenario(data, loan, obligor, flags)
        result = calculate_portfolio(augmented)
        scenario_asset = next(a for a in result.assets if a.obligor_name == "Test Holdco")
        self.assertTrue(scenario_asset.eligibility.eligible)

    def test_scenario_ineligible_pik(self):
        """Scenario loan with PIK should be ineligible (Test E)."""
        data = _make_data([])
        scenario = self._scenario(pik_pct="0.05")
        loan, obligor, flags = build_scenario_records(scenario)
        augmented = inject_scenario(data, loan, obligor, flags)
        result = calculate_portfolio(augmented)
        asset = next(a for a in result.assets if a.obligor_name == "Test Holdco")
        self.assertFalse(asset.eligibility.eligible)
        self.assertFalse(asset.eligibility.test_non_pik_dip)


# ---------------------------------------------------------------------------
# TestWaterfallUnit — synthetic edge cases for concentration waterfall + WAAR
# ---------------------------------------------------------------------------

class TestWaterfallUnit(unittest.TestCase):
    """Synthetic tests for _run_waterfall() edge cases.

    Each test constructs a minimal WorkbookData with known inputs and asserts
    on the waterfall outputs. No workbook required.
    """

    # ---- Zero-ABV portfolio ------------------------------------------------

    def test_zero_abv_waterfall_is_none(self):
        """Empty portfolio (no eligible loans) → waterfall is None, totals zero.

        calculate_portfolio skips _run_waterfall when eligible_assets is empty.
        This is the documented contract: callers must guard against waterfall=None.
        """
        data = _make_data([])
        calc = calculate_portfolio(data)
        self.assertIsNone(calc.waterfall)
        self.assertEqual(calc.total_pre_conc_eligible_value, 0.0)
        self.assertEqual(calc.total_borrowing_value, 0.0)
        self.assertEqual(len(calc.eligible_assets), 0)

    # ---- All-2L portfolio (Max Second Lien test) ---------------------------

    def test_all_second_lien_waterfall_sequence(self):
        """All-Second Lien portfolio: test 1a fires first, then test 1b sees haircutted values.

        Test 1a ("Max Second Lien & FILO with senior lev >= 1.50x") qualifies ALL
        Second Lien loans unconditionally — the >= 1.50x threshold applies only to FILO.
        So for a 100% Second Lien portfolio:
          - Test 1a: qual=total ABV, excess fired, haircut applied
          - Test 1b: same loans re-evaluated on post-haircut values → excess=0 (already capped)
        This documents the waterfall sequential haircut contract.
        """
        loans = [
            _make_loan(f"Co{i}", row=i, loan_type="Second Lien",
                       current_ebitda_mm=15.0, net_detachment=1.0, net_attachment=1.5,
                       olb=10_000_000)
            for i in range(1, 6)
        ]
        flags = {f"Co{i}": _make_flags(f"Co{i}") for i in range(1, 6)}
        data = _make_data(loans, flags=flags)
        calc = calculate_portfolio(data)
        w = calc.waterfall

        # Test 1a fires — all 5 Second Lien loans qualify unconditionally
        test_1a = next(t for t in w.concentration_tests
                       if t.limit_type == "Max Second Lien & FILO with senior lev >= 1.50x")
        self.assertGreater(test_1a.excess, 0.0)

        # Test 1b sees post-haircut values — excess already eliminated by test 1a
        test_1b = next(t for t in w.concentration_tests if t.limit_type == "Max Second Lien")
        self.assertEqual(test_1b.excess, 0.0)

        # Net ABV is reduced by test 1a haircut
        self.assertLess(w.net_abv, w.total_abv)

    # ---- Single-obligor portfolio (Max Obligors test) ----------------------

    def test_single_large_obligor_triggers_obligor_test(self):
        """One obligor whose value exceeds the per-obligor cap triggers haircut.

        Policy: max_obligor = 7.5%. With 1 obligor = 100% of ABV, excess = 92.5%.
        """
        loan = _make_loan("BigCo", row=1, olb=50_000_000)
        flags = {"BigCo": _make_flags("BigCo")}
        data = _make_data([loan], flags=flags)
        calc = calculate_portfolio(data)
        w = calc.waterfall

        obligor_test = next(t for t in w.concentration_tests if t.limit_type == "Max Obligors")
        self.assertGreater(obligor_test.excess, 0.0)

    # ---- WAAR cap by obligor count -----------------------------------------

    def test_waar_capped_at_low_diversity_below_12_obligors(self):
        """< 12 eligible obligors → WAAR capped at cap_low_diversity (0.50)."""
        # 3 eligible loans, all First Lien large EBITDA → uncapped WAAR = 0.675
        # cap_low_diversity = 0.50 → final_waar should be 0.50
        loans = [
            _make_loan(f"Co{i}", row=i, olb=10_000_000, current_ebitda_mm=25.0)
            for i in range(1, 4)
        ]
        flags = {f"Co{i}": _make_flags(f"Co{i}") for i in range(1, 4)}
        data = _make_data(loans, flags=flags)
        calc = calculate_portfolio(data)
        w = calc.waterfall

        self.assertLess(w.discrete_obligor_count, 12)
        self.assertAlmostEqual(w.applicable_waar_cap, 0.50)
        self.assertLessEqual(w.final_waar, 0.50 + 1e-9)

    def test_waar_not_capped_when_above_raw(self):
        """When raw_waar < cap, final_waar == raw_waar (cap not binding)."""
        # Second Lien loan → advance rate 0.55, which is below any cap (0.50, 0.60, 0.65)
        loan = _make_loan("Co1", row=1, loan_type="Second Lien",
                          current_ebitda_mm=15.0, net_detachment=2.0, olb=10_000_000)
        flags = {"Co1": _make_flags("Co1")}
        # Use high-diversity policy so cap doesn't apply
        policy = _make_policy(cap_high_diversity=0.65, threshold_high_diversity=1)
        data = _make_data([loan], flags=flags, policy=policy)
        calc = calculate_portfolio(data)
        w = calc.waterfall

        # raw_waar should be ~0.55 (2L rate), which is below any cap
        self.assertAlmostEqual(w.raw_waar, w.final_waar, places=4)

    # ---- Concentration test label integrity --------------------------------

    def test_all_13_concentration_labels_present(self):
        """Waterfall always returns exactly 13 tests with the canonical labels."""
        from borrowing_base_workbench.calculator import CONCENTRATION_TESTS
        loans = [_make_loan(f"Co{i}", row=i) for i in range(1, 4)]
        flags = {f"Co{i}": _make_flags(f"Co{i}") for i in range(1, 4)}
        data = _make_data(loans, flags=flags)
        calc = calculate_portfolio(data)

        result_labels = [t.limit_type for t in calc.waterfall.concentration_tests]
        expected_labels = [label for label, _ in CONCENTRATION_TESTS]
        self.assertEqual(result_labels, expected_labels)

    def test_select_active_vae_no_vaes_returns_none(self):
        """_select_active_vae with empty list → None."""
        from borrowing_base_workbench.calculator import _select_active_vae
        self.assertIsNone(_select_active_vae([]))

    def test_select_active_vae_all_null_values_returns_none(self):
        """_select_active_vae when all VAEs have None assigned_value → None."""
        from borrowing_base_workbench.calculator import _select_active_vae
        vae = VaeRecord(
            source_row=1, borrower="A", event_type="x", specific_test="",
            material_modification="", vae_date=date(2024, 1, 1),
            ebitda_at_vae=None, agent_adj_haircut=None, permitted_ebitda=None,
            senior_debt=None, total_debt=None, net_senior_leverage=None,
            net_total_leverage=None, interest_coverage_at_vae=None,
            vae_agent_assigned_value=None,
        )
        self.assertIsNone(_select_active_vae([vae]))

    def test_select_active_vae_null_date_loses_to_dated(self):
        """_select_active_vae: undated VAE loses to a dated one."""
        from borrowing_base_workbench.calculator import _select_active_vae
        def _vae(d, val):
            return VaeRecord(
                source_row=1, borrower="A", event_type="x", specific_test="",
                material_modification="", vae_date=d,
                ebitda_at_vae=None, agent_adj_haircut=None, permitted_ebitda=None,
                senior_debt=None, total_debt=None, net_senior_leverage=None,
                net_total_leverage=None, interest_coverage_at_vae=None,
                vae_agent_assigned_value=val,
            )
        undated = _vae(None, 9_000_000.0)
        dated   = _vae(date(2024, 6, 1), 3_000_000.0)
        # dated is more recent (None treated as date.min) → should win
        self.assertAlmostEqual(_select_active_vae([undated, dated]), 3_000_000.0)


# ---------------------------------------------------------------------------
# TestIntegration — full pipeline against real workbook
# ---------------------------------------------------------------------------

@unittest.skipUnless(WORKBOOK.exists(), f"Workbook not found at {WORKBOOK}")
class TestIntegration(unittest.TestCase):
    """End-to-end regression tests against the 2025-02-11 workbook.

    All reference values are from concentration_implementation_note.md and
    calculator_implementation_note.md, verified against workbook output.
    Tolerances:
      - Dollar amounts: ±$10 (floating-point accumulation allowance)
      - Rates: ±0.0001 (0.01%)
      - Counts: exact
    """

    TOL_DOLLARS = 10
    TOL_RATE = 0.0001

    @classmethod
    def setUpClass(cls):
        from borrowing_base_workbench.loader import load_workbook_data
        cls.data = load_workbook_data(WORKBOOK)
        cls.calc = calculate_portfolio(cls.data)
        cls.w = cls.calc.waterfall
        cls.av = cls.w.availability

    # ---- Asset counts ------------------------------------------------------

    def test_total_asset_count(self):
        self.assertEqual(len(self.calc.assets), REF["asset_count"])

    def test_eligible_count(self):
        self.assertEqual(len(self.calc.eligible_assets), REF["eligible_count"])

    def test_ineligible_count(self):
        self.assertEqual(len(self.calc.ineligible_assets), REF["ineligible_count"])

    def test_vae_affected_count(self):
        self.assertEqual(len(self.calc.vae_affected), REF["vae_affected_count"])

    # ---- Portfolio par ------------------------------------------------------

    def test_total_portfolio_par(self):
        self.assertAlmostEqual(
            self.calc.total_portfolio_par,
            REF["total_portfolio_par"],
            delta=self.TOL_DOLLARS,
        )

    # ---- Pre-concentration values ------------------------------------------

    def test_aggregate_adjusted_bv(self):
        self.assertAlmostEqual(
            self.calc.total_pre_conc_eligible_value,
            REF["total_pre_conc_eligible_value"],
            delta=self.TOL_DOLLARS,
        )

    def test_total_borrowing_value(self):
        self.assertAlmostEqual(
            self.calc.total_borrowing_value,
            REF["total_borrowing_value"],
            delta=self.TOL_DOLLARS,
        )

    # ---- WAAR and obligor count --------------------------------------------

    def test_raw_waar(self):
        self.assertAlmostEqual(self.w.raw_waar, REF["raw_waar"], delta=self.TOL_RATE)

    def test_discrete_obligor_count(self):
        self.assertEqual(self.w.discrete_obligor_count, REF["discrete_obligor_count"])

    def test_waar_cap_applied(self):
        self.assertAlmostEqual(self.w.applicable_waar_cap, REF["waar_cap"], delta=self.TOL_RATE)

    def test_final_waar_unconstrained(self):
        """With 25 obligors (≥20), cap = 65% which raw WAAR 60.80% doesn't hit."""
        self.assertAlmostEqual(self.w.final_waar, REF["final_waar"], delta=self.TOL_RATE)

    # ---- Concentration waterfall -------------------------------------------

    def test_total_excess_concentration(self):
        self.assertAlmostEqual(
            self.w.total_excess, REF["total_excess"], delta=self.TOL_DOLLARS
        )

    def test_net_abv(self):
        self.assertAlmostEqual(self.w.net_abv, REF["net_abv"], delta=self.TOL_DOLLARS)

    def test_obligor_concentration_excess(self):
        """Max Obligors test produces the largest excess ($24.9M)."""
        test = next(t for t in self.w.concentration_tests if t.limit_type == "Max Obligors")
        self.assertAlmostEqual(test.excess, REF["obligor_excess"], delta=self.TOL_DOLLARS)

    def test_industry_concentration_excess(self):
        """Max Largest Industry produces $4.4M excess."""
        test = next(
            t for t in self.w.concentration_tests if t.limit_type == "Max Largest Industry"
        )
        self.assertAlmostEqual(test.excess, REF["industry_excess"], delta=self.TOL_DOLLARS)

    def test_zero_excess_tests(self):
        """All tests except Max Obligors and Max Largest Industry should have zero excess."""
        non_zero = {
            t.limit_type for t in self.w.concentration_tests if t.excess > 0
        }
        self.assertSetEqual(non_zero, {"Max Obligors", "Max Largest Industry"})

    # ---- Availability ------------------------------------------------------

    def test_test_a_facility(self):
        self.assertAlmostEqual(
            self.av.test_a_facility, REF["test_a_facility"], delta=self.TOL_DOLLARS
        )

    def test_test_b_borrowing_base(self):
        self.assertAlmostEqual(
            self.av.test_b_borrowing_base, REF["test_b_borrowing_base"], delta=self.TOL_DOLLARS
        )

    def test_test_c_credit_enhancement(self):
        self.assertAlmostEqual(
            self.av.test_c_credit_enhancement, REF["test_c_credit_enhancement"], delta=self.TOL_DOLLARS
        )

    def test_availability(self):
        self.assertAlmostEqual(
            self.av.availability, REF["availability"], delta=self.TOL_DOLLARS
        )

    def test_test_b_is_binding(self):
        """Test B is the binding (lowest) constraint in the current workbook state."""
        self.assertLessEqual(self.av.test_b_borrowing_base, self.av.test_a_facility)
        self.assertLessEqual(self.av.test_b_borrowing_base, self.av.test_c_credit_enhancement)
        self.assertAlmostEqual(self.av.availability, self.av.test_b_borrowing_base, delta=1)

    # ---- Engine API contract -----------------------------------------------

    def test_probe_workbook_returns_ok(self):
        from borrowing_base_workbench.engine import probe_workbook
        result = probe_workbook(WORKBOOK)
        self.assertEqual(result["status"], "ok")
        self.assertIn("metrics", result)
        self.assertIn("concentration_limits", result)
        self.assertEqual(len(result["concentration_limits"]), 13)

    def test_probe_workbook_metrics_keys(self):
        from borrowing_base_workbench.engine import probe_workbook
        result = probe_workbook(WORKBOOK)
        required = {
            "availability", "aggregate_adjusted_bv", "excess_concentration",
            "net_adjusted_bv", "credit_enhancement_test",
            "weighted_avg_advance_rate", "current_advances",
        }
        self.assertTrue(required.issubset(result["metrics"].keys()))

    def test_probe_workbook_availability_value(self):
        from borrowing_base_workbench.engine import probe_workbook
        result = probe_workbook(WORKBOOK)
        self.assertAlmostEqual(
            result["metrics"]["availability"], REF["availability"], delta=self.TOL_DOLLARS
        )

    def test_probe_workbook_concentration_labels(self):
        from borrowing_base_workbench.engine import probe_workbook, _CONCENTRATION_LABEL_ORDER
        result = probe_workbook(WORKBOOK)
        returned_labels = [r["limit_type"] for r in result["concentration_limits"]]
        self.assertEqual(returned_labels, _CONCENTRATION_LABEL_ORDER)

    def test_run_pro_forma_contract(self):
        """run_pro_forma() must return {status, before, after, eligibility} with all keys."""
        from borrowing_base_workbench.engine import run_pro_forma
        scenario = {
            "company_name": "Integration Test Co",
            "security_type": "First Lien",
            "ltm_revenue": "80000000",
            "ltm_adj_ebitda": "20000000",
            "drawn_revolver": "0",
            "first_out_balance": "0",
            "pari_passu": "0",
            "bdc_balance": "15000000",
            "total_sm_balance": "15000000",
            "cash_balance": "0",
            "interest_coverage": "2.5",
            "loan_denomination": "USD",
            "purchase_price": "1.0",
            "country": "United States",
            "investment_date": "2025-03-01",
            "maturity_date": "2031-03-01",
            "rate_type": "Floating",
            "industry_classification": "Software",
            "payment_frequency": "M",
            "pik_pct": "0",
            "spread": "0.055",
            "sofr_floor": "0.005",
            "attach_point": "0.0",
        }
        result = run_pro_forma(WORKBOOK, scenario)
        self.assertEqual(result["status"], "ok")

        # Both snapshots present
        for snap_key in ("before", "after"):
            snap = result[snap_key]
            for key in ("availability", "aggregate_adjusted_bv", "excess_concentration",
                        "net_adjusted_bv", "weighted_avg_advance_rate",
                        "credit_enhancement_test", "current_advances"):
                self.assertIn(key, snap, f"Missing {key!r} from {snap_key!r}")
            self.assertIn("concentration_limits", snap)
            self.assertEqual(len(snap["concentration_limits"]), 13)

        # Eligibility result present
        elig = result["eligibility"]
        self.assertIn("status", elig)
        self.assertIn("failed_tests", elig)

    def test_run_pro_forma_before_matches_probe(self):
        """run_pro_forma() 'before' state must match probe_workbook() metrics."""
        from borrowing_base_workbench.engine import probe_workbook, run_pro_forma
        probe = probe_workbook(WORKBOOK)["metrics"]
        scenario = {
            "company_name": "X", "security_type": "First Lien",
            "ltm_revenue": "80000000", "ltm_adj_ebitda": "20000000",
            "drawn_revolver": "0", "first_out_balance": "0", "pari_passu": "0",
            "bdc_balance": "15000000", "total_sm_balance": "15000000",
            "cash_balance": "0", "interest_coverage": "2.5",
            "loan_denomination": "USD", "purchase_price": "1.0",
            "country": "United States", "investment_date": "2025-03-01",
            "maturity_date": "2031-03-01", "rate_type": "Floating",
            "industry_classification": "Software", "payment_frequency": "M",
            "pik_pct": "0", "spread": "0.055", "sofr_floor": "0.005", "attach_point": "0.0",
        }
        pf = run_pro_forma(WORKBOOK, scenario)["before"]
        for key in ("availability", "aggregate_adjusted_bv", "excess_concentration",
                    "net_adjusted_bv", "weighted_avg_advance_rate"):
            self.assertAlmostEqual(pf[key], probe[key], delta=self.TOL_DOLLARS,
                                   msg=f"Mismatch for {key!r}")

    def test_scenario_clean_eligible(self):
        """A well-formed First Lien scenario should be eligible."""
        from borrowing_base_workbench.engine import run_pro_forma
        scenario = {
            "company_name": "Clean Co", "security_type": "First Lien",
            "ltm_revenue": "100000000", "ltm_adj_ebitda": "25000000",
            "drawn_revolver": "0", "first_out_balance": "0", "pari_passu": "0",
            "bdc_balance": "20000000", "total_sm_balance": "20000000",
            "cash_balance": "0", "interest_coverage": "2.5",
            "loan_denomination": "USD", "purchase_price": "1.0",
            "country": "United States", "investment_date": "2025-03-01",
            "maturity_date": "2031-03-01", "rate_type": "Floating",
            "industry_classification": "Software", "payment_frequency": "M",
            "pik_pct": "0", "spread": "0.055", "sofr_floor": "0.005", "attach_point": "0.0",
        }
        result = run_pro_forma(WORKBOOK, scenario)
        self.assertEqual(result["eligibility"]["status"], "Yes")
        self.assertEqual(result["eligibility"]["failed_tests"], [])

    def test_scenario_pik_ineligible(self):
        """Scenario loan with PIK must produce eligibility status = No."""
        from borrowing_base_workbench.engine import run_pro_forma
        scenario = {
            "company_name": "PIK Co", "security_type": "First Lien",
            "ltm_revenue": "100000000", "ltm_adj_ebitda": "25000000",
            "drawn_revolver": "0", "first_out_balance": "0", "pari_passu": "0",
            "bdc_balance": "20000000", "total_sm_balance": "20000000",
            "cash_balance": "0", "interest_coverage": "2.5",
            "loan_denomination": "USD", "purchase_price": "1.0",
            "country": "United States", "investment_date": "2025-03-01",
            "maturity_date": "2031-03-01", "rate_type": "Floating",
            "industry_classification": "Software", "payment_frequency": "M",
            "pik_pct": "0.05",  # ← fails Test E
            "spread": "0.055", "sofr_floor": "0.005", "attach_point": "0.0",
        }
        result = run_pro_forma(WORKBOOK, scenario)
        self.assertEqual(result["eligibility"]["status"], "No")
        self.assertTrue(len(result["eligibility"]["failed_tests"]) >= 1)

    def test_scenario_adds_availability(self):
        """Adding a clean eligible loan must increase availability."""
        from borrowing_base_workbench.engine import run_pro_forma
        scenario = {
            "company_name": "Additive Co", "security_type": "First Lien",
            "ltm_revenue": "100000000", "ltm_adj_ebitda": "25000000",
            "drawn_revolver": "0", "first_out_balance": "0", "pari_passu": "0",
            "bdc_balance": "20000000", "total_sm_balance": "20000000",
            "cash_balance": "0", "interest_coverage": "2.5",
            "loan_denomination": "USD", "purchase_price": "1.0",
            "country": "United States", "investment_date": "2025-03-01",
            "maturity_date": "2031-03-01", "rate_type": "Floating",
            "industry_classification": "Software", "payment_frequency": "M",
            "pik_pct": "0", "spread": "0.055", "sofr_floor": "0.005", "attach_point": "0.0",
        }
        result = run_pro_forma(WORKBOOK, scenario)
        self.assertGreater(result["after"]["availability"], result["before"]["availability"])

    def test_scenario_from_json_file(self):
        """Scenario loaded from scratch/scenario_test.json should run without error."""
        import json
        scenario_path = WORKBOOK.parents[1] / "borrowing-base" / "scratch" / "scenario_test.json"
        if not scenario_path.exists():
            self.skipTest("scenario_test.json not found")
        from borrowing_base_workbench.engine import run_pro_forma
        scenario = json.loads(scenario_path.read_text(encoding="utf-8-sig"))
        result = run_pro_forma(WORKBOOK, scenario)
        self.assertEqual(result["status"], "ok")


# ---------------------------------------------------------------------------
# Entry point
# ---------------------------------------------------------------------------

if __name__ == "__main__":
    unittest.main(verbosity=2)
