"""calculator.py — Asset-level borrowing base calculation engine (Phase 3).

Computes per-loan eligibility, collateral tier, VAE-adjusted assigned value,
and advance rate bucket from loaded WorkbookData.

Phase 3 scope (this module):
  - Eligibility: 10 boolean tests per loan → eligible flag + failed reasons
  - Collateral tier: leverage-based tier lookup (4 tiers from PolicyConfig)
  - VAE override: assigned_value = min(fmv, vae_assigned, tier_value)
  - Advance rate: EBITDA/leverage-based bucketing

NOT in this phase (Phase 4+):
  - Concentration waterfall (13 sequential haircut tests)
  - Weighted-average advance rate cap (obligor count)
  - Availability calculation (MIN of 3 tests)
  - Scenario loan construction

Usage:
    from borrowing_base_workbench.calculator import calculate_portfolio
    result = calculate_portfolio(data)
    print(result.summary())
"""

from __future__ import annotations

from collections import Counter
from dataclasses import dataclass, field
from datetime import date
from pathlib import Path
from typing import Optional

from borrowing_base_workbench.loader import (
    CollateralTier,
    LoanRecord,
    ManualPortfolioFlags,
    PolicyConfig,
    VaeRecord,
    WorkbookData,
)


# ---------------------------------------------------------------------------
# Eligibility
# ---------------------------------------------------------------------------

# Loan types the facility agreement recognises as eligible
_VALID_LOAN_TYPES = {"First Lien", "FILO", "Second Lien", "DDTL", "Revolver"}

# Country strings treated as US domicile
_US_COUNTRY_VALUES = {"UNITED STATES", "US", "USA", "U.S.", "U.S.A."}

# VAE event-type keywords that indicate a bankruptcy / payment default
_BANKRUPTCY_KEYWORDS = {"bankrupt", "bankruptcy", "chapter 11", "chapter 7"}


@dataclass
class EligibilityResult:
    """10-test eligibility evaluation for one loan.

    Tests correspond to Portfolio cols E–N; master result mirrors col P.
    Each test_* field is True if the loan passes that test.
    failed_tests is a list of human-readable reasons for any failures.
    """

    # Test A: loan type is in the eligible-asset definition
    test_valid_type: bool
    # Test B: denominated in USD
    test_usd_denomination: bool
    # Test C: US-domiciled obligor
    test_us_domicile: bool
    # Test D: not covenant-lite
    test_not_cov_lite: bool
    # Test E: no PIK component AND not a DIP loan
    test_non_pik_dip: bool
    # Test F: acquisition price ≥ policy minimum (typically 90 cents)
    test_acquisition_price: bool
    # Test G: if Second Lien, obligor EBITDA ≥ $10MM; N/A (True) for other types
    test_second_lien_ebitda: bool
    # Test H: initial total leverage ≤ policy maximum (typically 6.5x)
    test_total_leverage: bool
    # Test I: original maturity ≤ policy maximum years, measured from inclusion_date
    test_original_maturity: bool
    # Test J: no active bankruptcy/default VAE event for this obligor
    test_not_defaulted: bool

    eligible: bool
    failed_tests: list[str]

    @classmethod
    def compute(
        cls,
        loan: LoanRecord,
        flags: ManualPortfolioFlags,
        policy: PolicyConfig,
        vaes: list[VaeRecord],
    ) -> "EligibilityResult":
        failed: list[str] = []

        # A: valid loan type
        t_type = loan.loan_type in _VALID_LOAN_TYPES
        if not t_type:
            failed.append(f"Invalid loan type: {loan.loan_type!r}")

        # B: USD denomination
        t_usd = loan.denomination.upper() in {"USD", "US DOLLAR"}
        if not t_usd:
            failed.append(f"Non-USD denomination: {loan.denomination!r}")

        # C: US domicile
        t_domicile = loan.country.upper() in _US_COUNTRY_VALUES
        if not t_domicile:
            failed.append(f"Non-US domicile: {loan.country!r}")

        # D: not cov-lite
        t_cov_lite = not loan.cov_lite
        if not t_cov_lite:
            failed.append("Covenant-lite loan")

        # E: non-PIK and non-DIP
        t_pik_dip = loan.pik_pct <= 0.0 and not flags.is_dip
        if not t_pik_dip:
            reasons: list[str] = []
            if loan.pik_pct > 0.0:
                reasons.append(f"PIK={loan.pik_pct:.1%}")
            if flags.is_dip:
                reasons.append("DIP")
            failed.append(f"PIK/DIP: {', '.join(reasons)}")

        # F: acquisition price ≥ threshold (e.g. 0.90 = 90 cents)
        t_price = loan.purchase_price >= policy.min_acquisition_price
        if not t_price:
            failed.append(
                f"Acquisition price {loan.purchase_price:.4f} < "
                f"minimum {policy.min_acquisition_price:.4f}"
            )

        # G: 2L EBITDA threshold (N/A if not Second Lien)
        if loan.loan_type == "Second Lien":
            # current_ebitda_mm is in millions; min_ebitda_second_lien_mm is in millions
            t_2l_ebitda = loan.current_ebitda_mm >= policy.min_ebitda_second_lien_mm
            if not t_2l_ebitda:
                failed.append(
                    f"2L EBITDA {loan.current_ebitda_mm:.1f}MM < "
                    f"{policy.min_ebitda_second_lien_mm:.1f}MM minimum"
                )
        else:
            t_2l_ebitda = True  # N/A for non-Second-Lien types

        # H: initial total leverage at inclusion ≤ max (None → pass; data absent for legacy loans)
        if loan.initial_total_leverage is None:
            t_leverage = True
        else:
            t_leverage = loan.initial_total_leverage <= policy.max_total_leverage_at_inclusion
            if not t_leverage:
                failed.append(
                    f"Initial total leverage {loan.initial_total_leverage:.2f}x > "
                    f"max {policy.max_total_leverage_at_inclusion:.2f}x"
                )

        # I: original maturity ≤ max years, measured from inclusion_date (col AO, NOT col U)
        if loan.maturity_date is None or loan.inclusion_date is None:
            t_maturity = False
            if loan.maturity_date is None:
                failed.append("Missing maturity date")
            if loan.inclusion_date is None:
                failed.append("Missing inclusion date")
        else:
            orig_years = (loan.maturity_date - loan.inclusion_date).days / 365.0
            t_maturity = orig_years <= policy.max_original_maturity_years
            if not t_maturity:
                failed.append(
                    f"Original maturity {orig_years:.2f}y > "
                    f"max {policy.max_original_maturity_years:.2f}y"
                )

        # J: no bankruptcy VAE event for this obligor
        has_bankruptcy = any(
            any(kw in (v.event_type or "").lower() for kw in _BANKRUPTCY_KEYWORDS)
            for v in vaes
        )
        t_not_defaulted = not has_bankruptcy
        if not t_not_defaulted:
            failed.append("Active bankruptcy/default VAE event")

        eligible = all([
            t_type, t_usd, t_domicile, t_cov_lite, t_pik_dip,
            t_price, t_2l_ebitda, t_leverage, t_maturity, t_not_defaulted,
        ])

        return cls(
            test_valid_type=t_type,
            test_usd_denomination=t_usd,
            test_us_domicile=t_domicile,
            test_not_cov_lite=t_cov_lite,
            test_non_pik_dip=t_pik_dip,
            test_acquisition_price=t_price,
            test_second_lien_ebitda=t_2l_ebitda,
            test_total_leverage=t_leverage,
            test_original_maturity=t_maturity,
            test_not_defaulted=t_not_defaulted,
            eligible=eligible,
            failed_tests=failed,
        )


# ---------------------------------------------------------------------------
# Collateral tier lookup
# ---------------------------------------------------------------------------

def _lookup_collateral_tier(
    loan: LoanRecord,
    tiers: list[CollateralTier],
) -> tuple[CollateralTier | None, float]:
    """Return (matched_tier, applicable_pct) for this loan.

    First Lien / DDTL / Revolver: compare loan.net_detachment against
                                   CollateralTier.max_first_lien_lev
    FILO / Second Lien:            compare loan.net_attachment against
                                   CollateralTier.max_filo_2l_lev

    Returns (None, 0.0) if the relevant leverage metric is unavailable.
    """
    is_senior = loan.loan_type in {"First Lien", "DDTL", "Revolver"}
    leverage = loan.net_detachment if is_senior else loan.net_attachment

    if leverage is None:
        return None, 0.0

    # Sort ascending by the relevant threshold so we walk Tier 1 → Tier 4
    sorted_tiers = sorted(
        tiers,
        key=lambda t: t.max_first_lien_lev if is_senior else t.max_filo_2l_lev,
    )

    for tier in sorted_tiers:
        threshold = tier.max_first_lien_lev if is_senior else tier.max_filo_2l_lev
        if leverage <= threshold:
            return tier, tier.applicable_pct

    # Leverage exceeds all tiers → sentinel (should be the 0% tier appended by loader)
    last = sorted_tiers[-1]
    return last, last.applicable_pct


def _tier_index(tier: CollateralTier | None, tiers: list[CollateralTier]) -> int | None:
    """Return 1-based tier index (100% = 1, 92.5% = 2, 85% = 3, 0% = 4)."""
    if tier is None:
        return None
    # Tiers sorted descending by pct: 1.0 → 0.925 → 0.85 → 0.0
    sorted_desc = sorted(tiers, key=lambda t: t.applicable_pct, reverse=True)
    for i, t in enumerate(sorted_desc, 1):
        if t is tier:
            return i
    return None


# ---------------------------------------------------------------------------
# Advance rate bucketing
# ---------------------------------------------------------------------------

def _pick_advance_rate(
    loan: LoanRecord,
    policy: PolicyConfig,
) -> tuple[str, float]:
    """Return (bucket_label, advance_rate) for this loan.

    First Lien / DDTL / Revolver: EBITDA-based (current_ebitda_mm, in millions)
    FILO:                          net_detachment-based
    Second Lien:                   fixed rate
    """
    lt = loan.loan_type

    if lt in {"First Lien", "DDTL", "Revolver"}:
        ebitda = loan.current_ebitda_mm  # in millions
        if ebitda > 20.0:
            return "First Lien >$20MM EBITDA", policy.rate_first_lien_large
        elif ebitda >= 10.0:
            return "First Lien $10-$20MM EBITDA", policy.rate_first_lien_mid
        else:
            return "First Lien <$10MM EBITDA", policy.rate_first_lien_small

    elif lt == "FILO":
        nd = loan.net_detachment
        if nd is None:
            # No leverage data → use most conservative rate
            return "FILO (leverage unavailable)", policy.rate_filo_high_lev
        elif nd < 0.75:
            return "FILO net det <0.75x", policy.rate_filo_low_lev
        elif nd <= 1.25:
            return "FILO net det 0.75–1.25x", policy.rate_filo_mid_lev
        else:
            return "FILO net det >1.25x", policy.rate_filo_high_lev

    elif lt == "Second Lien":
        return "Second Lien (fixed)", policy.rate_second_lien

    else:
        # Ineligible type — should be caught by eligibility Test A; fallback here
        return f"Unknown type ({lt})", 0.0


# ---------------------------------------------------------------------------
# Result dataclasses
# ---------------------------------------------------------------------------

@dataclass
class AssetCalcResult:
    """Full asset-level calculation result for one loan."""

    # --- Identity ---
    loan_row: int
    obligor_name: str
    loan_type: str
    ccm_id: str

    # --- Raw source values used in calculations ---
    olb: float
    fmv: float
    current_ebitda_mm: float
    net_detachment: float | None
    net_attachment: float | None
    initial_total_leverage: float | None
    purchase_price: float
    pik_pct: float
    cov_lite: bool
    is_dip: bool
    maturity_date: date | None
    inclusion_date: date | None
    original_maturity_years: float | None  # (maturity - inclusion) / 365

    # --- Eligibility ---
    eligibility: EligibilityResult

    # --- Collateral tier ---
    collateral_tier_index: int | None    # 1=100%, 2=92.5%, 3=85%, 4=0%
    collateral_tier_pct: float           # e.g. 1.0, 0.925, 0.85, 0.0
    collateral_tier_value: float         # tier_pct × olb

    # --- VAE ---
    has_vae: bool
    vae_agent_assigned_value: float | None  # most recent VAE col P value

    # --- Pre-concentration assigned value ---
    # min(fmv, vae_assigned [if applicable], tier_value)  — 0.0 if ineligible
    pre_conc_value: float

    # --- Advance rate ---
    advance_rate_label: str
    advance_rate: float        # 0.0 if ineligible

    # --- Borrowing value (pre-concentration) ---
    borrowing_value: float     # advance_rate × pre_conc_value


@dataclass
class PortfolioCalcResult:
    """Aggregate Phase 3 results for the full portfolio."""

    assets: list[AssetCalcResult]

    @property
    def eligible_assets(self) -> list[AssetCalcResult]:
        return [a for a in self.assets if a.eligibility.eligible]

    @property
    def ineligible_assets(self) -> list[AssetCalcResult]:
        return [a for a in self.assets if not a.eligibility.eligible]

    @property
    def vae_affected(self) -> list[AssetCalcResult]:
        """Eligible assets with at least one VAE (assigned value was adjusted)."""
        return [a for a in self.assets if a.has_vae and a.eligibility.eligible]

    @property
    def total_pre_conc_eligible_value(self) -> float:
        """Sum of pre-concentration assigned values for eligible assets."""
        return sum(a.pre_conc_value for a in self.eligible_assets)

    @property
    def total_portfolio_par(self) -> float:
        """Sum of OLB across all loans (eligible + ineligible)."""
        return sum(a.olb for a in self.assets)

    @property
    def total_borrowing_value(self) -> float:
        """Sum of (advance_rate × pre_conc_value) for eligible assets."""
        return sum(a.borrowing_value for a in self.eligible_assets)

    @property
    def implied_weighted_avg_advance_rate(self) -> float:
        """Weighted-average advance rate (pre-concentration, pre-cap).

        = total_borrowing_value / total_pre_conc_eligible_value
        Note: subject to obligor-count cap in Phase 4.
        """
        base = self.total_pre_conc_eligible_value
        return self.total_borrowing_value / base if base > 0 else 0.0

    def summary(self) -> dict:
        return {
            "asset_count": len(self.assets),
            "eligible_count": len(self.eligible_assets),
            "ineligible_count": len(self.ineligible_assets),
            "vae_affected_count": len(self.vae_affected),
            "total_portfolio_par": self.total_portfolio_par,
            "total_pre_conc_eligible_value": self.total_pre_conc_eligible_value,
            "total_borrowing_value": self.total_borrowing_value,
            "implied_weighted_avg_advance_rate": self.implied_weighted_avg_advance_rate,
        }


# ---------------------------------------------------------------------------
# Core computation functions
# ---------------------------------------------------------------------------

def calculate_asset(loan: LoanRecord, data: WorkbookData) -> AssetCalcResult:
    """Compute full asset-level borrowing base calculation for one loan."""
    policy = data.policy
    flags = data.flags_for(loan.obligor_name)
    vaes = data.vaes_for(loan.obligor_name)

    # --- Eligibility ---
    elig = EligibilityResult.compute(loan, flags, policy, vaes)

    # --- Collateral tier ---
    tier, tier_pct = _lookup_collateral_tier(loan, policy.collateral_tiers)
    tier_idx = _tier_index(tier, policy.collateral_tiers)
    collateral_tier_value = tier_pct * loan.olb

    # --- VAE: find most recent record with a non-None assigned value ---
    vaes_with_value = [v for v in vaes if v.vae_agent_assigned_value is not None]
    if vaes_with_value:
        most_recent = max(vaes_with_value, key=lambda v: v.vae_date or date.min)
        vae_assigned = most_recent.vae_agent_assigned_value
    else:
        vae_assigned = None
    has_vae = bool(vaes)

    # --- Pre-concentration assigned value ---
    if not elig.eligible:
        pre_conc_value = 0.0
    else:
        candidates = [loan.fmv, collateral_tier_value]
        if vae_assigned is not None:
            candidates.append(vae_assigned)
        pre_conc_value = min(candidates)

    # --- Advance rate ---
    if elig.eligible:
        rate_label, rate = _pick_advance_rate(loan, policy)
    else:
        rate_label, rate = "N/A (ineligible)", 0.0

    borrowing_value = rate * pre_conc_value

    # --- Original maturity in years ---
    if loan.maturity_date and loan.inclusion_date:
        orig_maturity_years = (loan.maturity_date - loan.inclusion_date).days / 365.0
    else:
        orig_maturity_years = None

    return AssetCalcResult(
        loan_row=loan.source_row,
        obligor_name=loan.obligor_name,
        loan_type=loan.loan_type,
        ccm_id=loan.ccm_id,
        olb=loan.olb,
        fmv=loan.fmv,
        current_ebitda_mm=loan.current_ebitda_mm,
        net_detachment=loan.net_detachment,
        net_attachment=loan.net_attachment,
        initial_total_leverage=loan.initial_total_leverage,
        purchase_price=loan.purchase_price,
        pik_pct=loan.pik_pct,
        cov_lite=loan.cov_lite,
        is_dip=flags.is_dip,
        maturity_date=loan.maturity_date,
        inclusion_date=loan.inclusion_date,
        original_maturity_years=orig_maturity_years,
        eligibility=elig,
        collateral_tier_index=tier_idx,
        collateral_tier_pct=tier_pct,
        collateral_tier_value=collateral_tier_value,
        has_vae=has_vae,
        vae_agent_assigned_value=vae_assigned,
        pre_conc_value=pre_conc_value,
        advance_rate_label=rate_label,
        advance_rate=rate,
        borrowing_value=borrowing_value,
    )


def calculate_portfolio(data: WorkbookData) -> PortfolioCalcResult:
    """Run asset-level calculations for all loans in the portfolio."""
    assets = [calculate_asset(loan, data) for loan in data.loans]
    return PortfolioCalcResult(assets=assets)


# ---------------------------------------------------------------------------
# CLI debug / test path
# ---------------------------------------------------------------------------

if __name__ == "__main__":
    import sys

    wb_path = Path(r"C:\Users\henry.yan\Downloads\2025-02-11_BDC Borrowing_Base_v8.xlsm")
    if len(sys.argv) > 1:
        wb_path = Path(sys.argv[1])

    print(f"Loading: {wb_path}")
    from borrowing_base_workbench.loader import load_workbook_data

    data = load_workbook_data(wb_path)
    result = calculate_portfolio(data)
    s = result.summary()

    print(f"\n{'='*65}")
    print("  PHASE 3 — ASSET-LEVEL RESULTS  (pre-concentration)")
    print(f"{'='*65}")
    print(f"  Total assets:                  {s['asset_count']}")
    print(f"  Eligible:                      {s['eligible_count']}")
    print(f"  Ineligible:                    {s['ineligible_count']}")
    print(f"  VAE-affected (eligible):       {s['vae_affected_count']}")
    print(f"  Total portfolio par (OLB):     ${s['total_portfolio_par']:>14,.0f}")
    print(f"  Pre-conc eligible value:       ${s['total_pre_conc_eligible_value']:>14,.0f}")
    print(f"  Pre-conc borrowing value:      ${s['total_borrowing_value']:>14,.0f}")
    rate_pct = s['implied_weighted_avg_advance_rate'] * 100
    print(f"  Implied WAAR (pre-cap):        {rate_pct:.2f}%")
    print(f"{'='*65}")

    ineligible = result.ineligible_assets
    if ineligible:
        print(f"\nIneligible assets ({len(ineligible)}):")
        for a in ineligible:
            reasons = "; ".join(a.eligibility.failed_tests)
            print(f"  [{a.loan_row}] {a.obligor_name} ({a.loan_type}): {reasons}")

    vae_eligible = result.vae_affected
    if vae_eligible:
        print(f"\nVAE-affected eligible assets ({len(vae_eligible)}):")
        for a in vae_eligible:
            print(
                f"  [{a.loan_row}] {a.obligor_name}: "
                f"fmv=${a.fmv:,.0f}  vae_assigned=${a.vae_agent_assigned_value:,.0f}  "
                f"tier_val=${a.collateral_tier_value:,.0f}  "
                f"pre_conc=${a.pre_conc_value:,.0f}"
            )

    print("\nAdvance rate distribution (eligible loans):")
    dist = Counter(a.advance_rate_label for a in result.eligible_assets)
    for label, count in sorted(dist.items()):
        pool = sum(a.pre_conc_value for a in result.eligible_assets if a.advance_rate_label == label)
        print(f"  {count:2d}  {label}: ${pool:,.0f} pre-conc value")

    print("\nCollateral tier distribution (eligible loans):")
    tier_dist = Counter(a.collateral_tier_index for a in result.eligible_assets)
    tier_labels = {1: "100%", 2: "92.5%", 3: "85%", 4: "0% (ineligible tier)", None: "N/A"}
    for idx in sorted(tier_dist, key=lambda x: (x is None, x)):
        print(f"  Tier {idx} ({tier_labels.get(idx, '?')}): {tier_dist[idx]} loans")
