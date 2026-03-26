"""calculator.py — Asset-level and portfolio-level borrowing base calculation engine.

Computes per-loan eligibility, collateral tier, VAE-adjusted assigned value,
advance rate bucketing, the 13-test concentration waterfall, WAAR with obligor-count
cap, and the three-test availability calculation.

Phase 3 (asset-level — unchanged):
  - Eligibility: 10 boolean tests per loan → eligible flag + failed reasons
  - Collateral tier: leverage-based tier lookup (4 tiers from PolicyConfig)
  - VAE override: assigned_value = min(fmv, vae_assigned, tier_value)
  - Advance rate: EBITDA/leverage-based bucketing

Phase 4 (added here — concentration waterfall + availability):
  - Sequential concentration waterfall: 13 tests, each haircuts qualifying assets
  - WAAR: SUMPRODUCT(rates, pct_of_abv), capped by discrete obligor count
  - Availability: MIN(facility, net_abv × WAAR + collections, total_abv - CE_min + collections)

NOT yet implemented:
  - Scenario loan construction (Phase 5)
  - Unfunded exposure equity deduction from availability (stub = 0)

Usage:
    from borrowing_base_workbench.calculator import calculate_portfolio
    result = calculate_portfolio(data)
    print(result.summary())
"""

from __future__ import annotations

from collections import Counter, defaultdict
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

    # --- Concentration-test classification flags ---
    # Stored here so the waterfall doesn't need to reach back to WorkbookData.
    industry: str              # LoanRecord.industry (S&P classification string)
    is_fixed_rate: bool        # rate_type == "Fixed"
    limited_industry: bool     # ManualPortfolioFlags.limited_industry (col BV)
    non_sponsor: bool          # ManualPortfolioFlags.non_sponsor (col BW)
    div_recap: bool            # ManualPortfolioFlags.div_recap (col BX)


# ---------------------------------------------------------------------------
# Concentration waterfall result dataclasses (Phase 4)
# ---------------------------------------------------------------------------

@dataclass
class ConcentrationTestResult:
    """Result of one step in the sequential concentration waterfall.

    qualifying_value: sum of current (already-haircut) values for loans that
                      qualify for this test, at the moment this test runs.
    applicable_limit: limit_percent × total_abv (fixed; denominator never changes).
    excess:           max(0, qualifying_value - applicable_limit).
    """
    limit_type: str
    limit_percent: float
    qualifying_value: float
    applicable_limit: float
    excess: float


@dataclass
class AvailabilityResult:
    """Workbook Availability sheet: three tests, then MIN.

    Mirrors L22 / L23 / L25 / L26.

    Approximations vs workbook (noted in implementation note):
      - unfunded_equity (L37): set to 0
      - interest_reserve (L19): set to 0
    """
    test_a_facility: float           # L22 = facility_amount
    net_abv: float                   # L35 = total_abv - total_excess
    final_waar: float                # L48 = MIN(raw_waar, obligor_count_cap)
    available_collections: float     # L18 (cash in Principal Collection Account)
    test_b_borrowing_base: float     # L23 = net_abv × final_waar + collections
    total_abv: float                 # L33 = total pre-conc eligible value
    credit_enhancement_min: float    # L41 = MAX(15M, top-3-obligor sum)
    test_c_credit_enhancement: float # L25 = total_abv - CE_min + collections
    availability: float              # L26 = MIN(test_a, test_b, test_c)


@dataclass
class WaterfallResult:
    """Full Phase 4 output: concentration waterfall + WAAR cap + availability."""
    total_abv: float                           # aggregate assigned value (M43 in CL sheet)
    concentration_tests: list[ConcentrationTestResult]
    total_excess: float                        # M44 = sum of all test excesses
    net_abv: float                             # M45 = total_abv - total_excess

    # WAAR
    total_borrowing_value: float               # sum(advance_rate × pre_conc_value)
    raw_waar: float                            # G38 in CL = total_bv / total_abv
    discrete_obligor_count: int                # G34 in CL
    applicable_waar_cap: float                 # policy cap driven by obligor count
    final_waar: float                          # L48 = MIN(raw_waar, cap)

    # Availability
    availability: AvailabilityResult


@dataclass
class PortfolioCalcResult:
    """Aggregate calculation results for the full portfolio (Phases 3 + 4)."""

    assets: list[AssetCalcResult]
    waterfall: WaterfallResult | None = None   # populated by calculate_portfolio()

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
        return sum(a.pre_conc_value for a in self.eligible_assets)

    @property
    def total_portfolio_par(self) -> float:
        return sum(a.olb for a in self.assets)

    @property
    def total_borrowing_value(self) -> float:
        return sum(a.borrowing_value for a in self.eligible_assets)

    @property
    def implied_weighted_avg_advance_rate(self) -> float:
        base = self.total_pre_conc_eligible_value
        return self.total_borrowing_value / base if base > 0 else 0.0

    # Phase 4 convenience accessors (None-safe)
    @property
    def availability(self) -> float:
        return self.waterfall.availability.availability if self.waterfall else 0.0

    @property
    def net_abv(self) -> float:
        return self.waterfall.net_abv if self.waterfall else 0.0

    @property
    def total_excess(self) -> float:
        return self.waterfall.total_excess if self.waterfall else 0.0

    @property
    def final_waar(self) -> float:
        return self.waterfall.final_waar if self.waterfall else 0.0

    def summary(self) -> dict:
        w = self.waterfall
        return {
            "asset_count": len(self.assets),
            "eligible_count": len(self.eligible_assets),
            "ineligible_count": len(self.ineligible_assets),
            "vae_affected_count": len(self.vae_affected),
            "total_portfolio_par": self.total_portfolio_par,
            "total_pre_conc_eligible_value": self.total_pre_conc_eligible_value,
            "total_borrowing_value": self.total_borrowing_value,
            # Phase 4 fields (0 if waterfall not run)
            "total_excess": w.total_excess if w else 0.0,
            "net_abv": w.net_abv if w else 0.0,
            "final_waar": w.final_waar if w else self.implied_weighted_avg_advance_rate,
            "availability": w.availability.availability if w else 0.0,
            "discrete_obligor_count": w.discrete_obligor_count if w else 0,
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
        # Concentration classification flags
        industry=loan.industry,
        is_fixed_rate=loan.rate_type.upper() == "FIXED",
        limited_industry=flags.limited_industry,
        non_sponsor=flags.non_sponsor,
        div_recap=flags.div_recap,
    )


def _run_waterfall(
    eligible: list[AssetCalcResult],
    policy,
    avail_meta,
) -> WaterfallResult:
    """Run the 13-test sequential concentration waterfall + WAAR cap + availability.

    Waterfall sequencing
    --------------------
    Each test operates on a `cur_val` dict keyed by loan_row.  A qualifying
    excess is distributed proportionally across qualifying loans, reducing
    their cur_val before the next test runs.  The limit denominator (total_abv)
    never changes — it is the original sum of pre_conc_values (mirrors CL!M43
    = Portfolio!V61 which is a fixed reference cell in every test formula).

    Availability approximations (noted in implementation note):
      - L37 unfunded exposure equity deduction: treated as 0
      - L19 interest reserve deduction: treated as 0
    """
    total_abv = sum(a.pre_conc_value for a in eligible)
    total_borrowing_value = sum(a.borrowing_value for a in eligible)

    if total_abv <= 0:
        zero_tests = [
            ConcentrationTestResult(name, pct, 0.0, 0.0, 0.0)
            for name, pct in [
                ("Max Second Lien & FILO with senior lev >= 1.50x", policy.conc_max_2l_filo_high_lev),
                ("Max Second Lien",                                  policy.conc_max_second_lien),
                ("Max Non-First Lien",                               policy.conc_max_non_first_lien),
                ("Max EBITDA < $5MM",                                policy.conc_max_small_ebitda),
                ("Max Obligors",                                     policy.conc_max_obligor),
                ("Max Largest Industry",                             policy.conc_max_largest_industry),
                ("Max Second Largest Industry",                      policy.conc_max_second_industry),
                ("Max Other Industries",                             policy.conc_max_other_industries),
                ("Fixed Rate",                                       policy.conc_fixed_rate),
                ("Max Limited Industry",                             policy.conc_limited_industry),
                ("Max DDTL and Revolver",                            policy.conc_max_ddtl_revolver),
                ("Max Non-Sponsor/Non-Family Office",                policy.conc_max_non_sponsor),
                ("Max Div Recap Non-Sponsor/Non-Family Office",      policy.conc_max_div_recap),
            ]
        ]
        avail_zero = AvailabilityResult(
            test_a_facility=avail_meta.facility_amount,
            net_abv=0.0, final_waar=0.0,
            available_collections=avail_meta.available_collections,
            test_b_borrowing_base=avail_meta.available_collections,
            total_abv=0.0, credit_enhancement_min=15_000_000.0,
            test_c_credit_enhancement=avail_meta.available_collections - 15_000_000.0,
            availability=avail_meta.available_collections - 15_000_000.0,
        )
        return WaterfallResult(
            total_abv=0.0, concentration_tests=zero_tests,
            total_excess=0.0, net_abv=0.0,
            total_borrowing_value=0.0, raw_waar=0.0,
            discrete_obligor_count=0, applicable_waar_cap=policy.cap_low_diversity,
            final_waar=0.0, availability=avail_zero,
        )

    # Working values per asset; haircut in-place as tests run
    cur_val: dict[int, float] = {a.loan_row: a.pre_conc_value for a in eligible}
    tests: list[ConcentrationTestResult] = []

    def _apply(label: str, limit_pct: float, rows: set[int]) -> None:
        """Compute excess for a qualifying row-set and apply proportional haircut."""
        qual = [a for a in eligible if a.loan_row in rows]
        qual_val = sum(cur_val[a.loan_row] for a in qual)
        limit = limit_pct * total_abv
        excess = max(0.0, qual_val - limit)
        if excess > 0.0 and qual_val > 0.0:
            for a in qual:
                cur_val[a.loan_row] -= excess * cur_val[a.loan_row] / qual_val
        tests.append(ConcentrationTestResult(label, limit_pct, qual_val, limit, excess))

    # ------------------------------------------------------------------
    # Test 1a: Max Second Lien & FILO with senior net lev ≥ 1.50x
    # ------------------------------------------------------------------
    _apply(
        "Max Second Lien & FILO with senior lev >= 1.50x",
        policy.conc_max_2l_filo_high_lev,
        {a.loan_row for a in eligible
         if a.loan_type == "Second Lien"
         or (a.loan_type == "FILO" and (a.net_detachment or 0.0) >= 1.50)},
    )

    # ------------------------------------------------------------------
    # Test 1b: Max Second Lien
    # ------------------------------------------------------------------
    _apply(
        "Max Second Lien",
        policy.conc_max_second_lien,
        {a.loan_row for a in eligible if a.loan_type == "Second Lien"},
    )

    # ------------------------------------------------------------------
    # Test 2: Max Non-First Lien (Second Lien + FILO with senior lev > 1.00x)
    # ------------------------------------------------------------------
    _apply(
        "Max Non-First Lien",
        policy.conc_max_non_first_lien,
        {a.loan_row for a in eligible
         if a.loan_type == "Second Lien"
         or (a.loan_type == "FILO" and (a.net_detachment or 0.0) > 1.00)},
    )

    # ------------------------------------------------------------------
    # Test 3: Max EBITDA < $5MM
    # ------------------------------------------------------------------
    _apply(
        "Max EBITDA < $5MM",
        policy.conc_max_small_ebitda,
        {a.loan_row for a in eligible if a.current_ebitda_mm < 5.0},
    )

    # ------------------------------------------------------------------
    # Test 4: Max Obligors — per-obligor cap, sum of excesses
    # Each obligor's total value is capped at conc_max_obligor × total_abv.
    # ------------------------------------------------------------------
    obligor_limit = policy.conc_max_obligor * total_abv
    obligor_groups: dict[str, list[AssetCalcResult]] = defaultdict(list)
    for a in eligible:
        obligor_groups[a.obligor_name].append(a)

    obligor_total_excess = 0.0
    for obl_assets in obligor_groups.values():
        obl_val = sum(cur_val[a.loan_row] for a in obl_assets)
        obl_excess = max(0.0, obl_val - obligor_limit)
        if obl_excess > 0.0 and obl_val > 0.0:
            obligor_total_excess += obl_excess
            for a in obl_assets:
                cur_val[a.loan_row] -= obl_excess * cur_val[a.loan_row] / obl_val

    # "actual" for the obligor test = total_abv (every obligor contributes;
    # mirrors CL!L57 / Portfolio!FI61 semantics)
    tests.append(ConcentrationTestResult(
        "Max Obligors",
        policy.conc_max_obligor,
        total_abv,          # qualifying_value = full portfolio (mirrors FI61)
        obligor_limit,
        obligor_total_excess,
    ))

    # ------------------------------------------------------------------
    # Tests 5a / 5b / 5c: Industry concentration
    # Industries ranked by CURRENT values (post tests 1–4).
    # ------------------------------------------------------------------

    def _industry_vals() -> dict[str, float]:
        iv: dict[str, float] = defaultdict(float)
        for a in eligible:
            iv[a.industry] += cur_val[a.loan_row]
        return iv

    # 5a: Largest industry
    iv5a = _industry_vals()
    if iv5a:
        ranked = sorted(iv5a.items(), key=lambda x: x[1], reverse=True)
        top1_ind, top1_val = ranked[0]
        limit_5a = policy.conc_max_largest_industry * total_abv
        excess_5a = max(0.0, top1_val - limit_5a)
        if excess_5a > 0.0 and top1_val > 0.0:
            for a in eligible:
                if a.industry == top1_ind:
                    cur_val[a.loan_row] -= excess_5a * cur_val[a.loan_row] / top1_val
        tests.append(ConcentrationTestResult("Max Largest Industry", policy.conc_max_largest_industry, top1_val, limit_5a, excess_5a))
    else:
        tests.append(ConcentrationTestResult("Max Largest Industry", policy.conc_max_largest_industry, 0.0, policy.conc_max_largest_industry * total_abv, 0.0))
        top1_ind = ""

    # 5b: Second largest industry (re-rank after 5a haircut)
    iv5b = _industry_vals()
    if iv5b:
        ranked5b = sorted(iv5b.items(), key=lambda x: x[1], reverse=True)
        top2_inds = {r[0] for r in ranked5b[:2]}
        if len(ranked5b) >= 2:
            sec_ind, sec_val = ranked5b[1]
            limit_5b = policy.conc_max_second_industry * total_abv
            excess_5b = max(0.0, sec_val - limit_5b)
            if excess_5b > 0.0 and sec_val > 0.0:
                for a in eligible:
                    if a.industry == sec_ind:
                        cur_val[a.loan_row] -= excess_5b * cur_val[a.loan_row] / sec_val
            tests.append(ConcentrationTestResult("Max Second Largest Industry", policy.conc_max_second_industry, sec_val, limit_5b, excess_5b))
        else:
            tests.append(ConcentrationTestResult("Max Second Largest Industry", policy.conc_max_second_industry, 0.0, policy.conc_max_second_industry * total_abv, 0.0))
    else:
        top2_inds: set[str] = set()
        tests.append(ConcentrationTestResult("Max Second Largest Industry", policy.conc_max_second_industry, 0.0, policy.conc_max_second_industry * total_abv, 0.0))

    # 5c: Other industries — each capped individually at conc_max_other_industries
    iv5c: dict[str, float] = defaultdict(float)
    for a in eligible:
        if a.industry not in top2_inds:
            iv5c[a.industry] += cur_val[a.loan_row]

    limit_5c_per = policy.conc_max_other_industries * total_abv
    other_total_val = sum(iv5c.values())
    other_total_excess = 0.0
    for ind, ind_val in iv5c.items():
        ind_excess = max(0.0, ind_val - limit_5c_per)
        if ind_excess > 0.0 and ind_val > 0.0:
            other_total_excess += ind_excess
            for a in eligible:
                if a.industry == ind:
                    cur_val[a.loan_row] -= ind_excess * cur_val[a.loan_row] / ind_val
    tests.append(ConcentrationTestResult("Max Other Industries", policy.conc_max_other_industries, other_total_val, limit_5c_per, other_total_excess))

    # ------------------------------------------------------------------
    # Test 6: Fixed Rate
    # ------------------------------------------------------------------
    _apply(
        "Fixed Rate",
        policy.conc_fixed_rate,
        {a.loan_row for a in eligible if a.is_fixed_rate},
    )

    # ------------------------------------------------------------------
    # Test 7: Max Limited Industry (energy, travel, hospitality, leisure)
    # Uses ManualPortfolioFlags.limited_industry (Portfolio col BV).
    # ------------------------------------------------------------------
    limited_ind_names = set(policy.limited_industries)  # lowercase strings from AGENT
    _apply(
        "Max Limited Industry",
        policy.conc_limited_industry,
        {a.loan_row for a in eligible
         if a.limited_industry or a.industry.lower() in limited_ind_names},
    )

    # ------------------------------------------------------------------
    # Test 8: Max DDTL and Revolver
    # ------------------------------------------------------------------
    _apply(
        "Max DDTL and Revolver",
        policy.conc_max_ddtl_revolver,
        {a.loan_row for a in eligible if a.loan_type in {"DDTL", "Revolver"}},
    )

    # ------------------------------------------------------------------
    # Test 9a: Max Non-Sponsor/Non-Family Office
    # ------------------------------------------------------------------
    _apply(
        "Max Non-Sponsor/Non-Family Office",
        policy.conc_max_non_sponsor,
        {a.loan_row for a in eligible if a.non_sponsor},
    )

    # ------------------------------------------------------------------
    # Test 9b: Max Div Recap Non-Sponsor/Non-Family Office
    # ------------------------------------------------------------------
    _apply(
        "Max Div Recap Non-Sponsor/Non-Family Office",
        policy.conc_max_div_recap,
        {a.loan_row for a in eligible if a.div_recap},
    )

    # ------------------------------------------------------------------
    # Totals
    # ------------------------------------------------------------------
    total_excess = sum(t.excess for t in tests)
    net_abv = total_abv - total_excess

    # ------------------------------------------------------------------
    # WAAR + obligor-count cap
    # G38 = SUMPRODUCT(rates, pct_of_abv) = total_bv / total_abv
    # L48 = MIN(G38, IFS(obligor_count → cap))
    # ------------------------------------------------------------------
    raw_waar = total_borrowing_value / total_abv if total_abv > 0 else 0.0

    discrete_count = len({a.obligor_name for a in eligible})
    if discrete_count >= policy.threshold_high_diversity:
        waar_cap = policy.cap_high_diversity
    elif discrete_count >= policy.threshold_mid_diversity:
        waar_cap = policy.cap_mid_diversity
    else:
        waar_cap = policy.cap_low_diversity
    final_waar = min(raw_waar, waar_cap)

    # ------------------------------------------------------------------
    # Availability (Availability sheet L22 / L23 / L25 / L26)
    # ------------------------------------------------------------------
    facility    = avail_meta.facility_amount
    collections = avail_meta.available_collections

    # L41 = MAX($15M, top-3-obligor sum)  [Portfolio!FJ col = per-obligor ABV]
    obligor_abv: dict[str, float] = defaultdict(float)
    for a in eligible:
        obligor_abv[a.obligor_name] += a.pre_conc_value   # use original (L41 in workbook uses V col)
    top3 = sum(sorted(obligor_abv.values(), reverse=True)[:3])
    ce_min = max(15_000_000.0, top3)

    test_a = facility                                # L22
    test_b = net_abv * final_waar + collections      # L23 (L37=0, L19=0)
    test_c = total_abv - ce_min + collections        # L25 (L37=0, L19=0)
    availability = min(test_a, test_b, test_c)

    avail_result = AvailabilityResult(
        test_a_facility=test_a,
        net_abv=net_abv,
        final_waar=final_waar,
        available_collections=collections,
        test_b_borrowing_base=test_b,
        total_abv=total_abv,
        credit_enhancement_min=ce_min,
        test_c_credit_enhancement=test_c,
        availability=availability,
    )

    return WaterfallResult(
        total_abv=total_abv,
        concentration_tests=tests,
        total_excess=total_excess,
        net_abv=net_abv,
        total_borrowing_value=total_borrowing_value,
        raw_waar=raw_waar,
        discrete_obligor_count=discrete_count,
        applicable_waar_cap=waar_cap,
        final_waar=final_waar,
        availability=avail_result,
    )


def calculate_portfolio(data: WorkbookData) -> PortfolioCalcResult:
    """Run full portfolio calculation: asset-level (Phase 3) + waterfall (Phase 4)."""
    assets = [calculate_asset(loan, data) for loan in data.loans]
    result = PortfolioCalcResult(assets=assets)
    eligible = result.eligible_assets
    if eligible:
        result.waterfall = _run_waterfall(eligible, data.policy, data.availability_meta)
    return result


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
    w = result.waterfall

    SEP = "=" * 70

    # ---- Asset-level summary ----
    print(f"\n{SEP}")
    print("  PHASES 3+4 — FULL BORROWING BASE CALCULATION")
    print(SEP)
    print(f"  Total assets:                  {s['asset_count']}")
    print(f"  Eligible:                      {s['eligible_count']}")
    print(f"  Ineligible:                    {s['ineligible_count']}")
    print(f"  VAE-affected (eligible):       {s['vae_affected_count']}")
    print(f"  Total portfolio par (OLB):     ${s['total_portfolio_par']:>16,.0f}")
    print(f"  Pre-conc eligible value (ABV): ${s['total_pre_conc_eligible_value']:>16,.0f}")
    print(f"  Pre-conc borrowing value:      ${s['total_borrowing_value']:>16,.0f}")
    print()

    # ---- Waterfall summary ----
    if w:
        av = w.availability
        print(f"  Total excess concentration:    ${w.total_excess:>16,.0f}")
        print(f"  Net ABV (after haircuts):      ${w.net_abv:>16,.0f}")
        print(f"  Raw WAAR (pre-cap):            {w.raw_waar*100:>15.2f}%")
        print(f"  Discrete obligors:             {w.discrete_obligor_count:>16d}")
        print(f"  WAAR cap applied:              {w.applicable_waar_cap*100:>15.2f}%")
        print(f"  Final WAAR:                    {w.final_waar*100:>15.2f}%")
        print()
        print(f"  Test A  (facility max):        ${av.test_a_facility:>16,.0f}")
        print(f"  Test B  (net_abv × WAAR + cash):${av.test_b_borrowing_base:>15,.0f}")
        print(f"  Test C  (credit enhancement):  ${av.test_c_credit_enhancement:>16,.0f}")
        print(f"  *** AVAILABILITY:              ${av.availability:>16,.0f} ***")
        print(f"  Current advances:              ${data.availability_meta.current_advances:>16,.0f}")
        cushion = av.availability - data.availability_meta.current_advances
        print(f"  Available cushion:             ${cushion:>16,.0f}")

    print(SEP)

    # ---- Ineligible assets ----
    ineligible = result.ineligible_assets
    if ineligible:
        print(f"\nIneligible assets ({len(ineligible)}):")
        for a in ineligible:
            print(f"  [{a.loan_row}] {a.obligor_name} ({a.loan_type}): {'; '.join(a.eligibility.failed_tests)}")

    # ---- Concentration test detail ----
    if w:
        print(f"\n{'Concentration tests':}")
        hdr = f"  {'Test':<50} {'Actual':>14}  {'Limit':>14}  {'Excess':>12}"
        print(hdr)
        print("  " + "-" * 95)
        for t in w.concentration_tests:
            flag = " <<<" if t.excess > 0 else ""
            print(
                f"  {t.limit_type:<50} "
                f"${t.qualifying_value:>13,.0f}  "
                f"${t.applicable_limit:>13,.0f}  "
                f"${t.excess:>11,.0f}{flag}"
            )
        print("  " + "-" * 95)
        print(f"  {'TOTAL EXCESS':<50}  {'':>14}  {'':>14}  ${w.total_excess:>11,.0f}")

    # ---- Advance rate distribution ----
    print("\nAdvance rate distribution (eligible loans):")
    dist = Counter(a.advance_rate_label for a in result.eligible_assets)
    for label, count in sorted(dist.items()):
        pool = sum(a.pre_conc_value for a in result.eligible_assets if a.advance_rate_label == label)
        print(f"  {count:2d}  {label}: ${pool:,.0f} pre-conc value")
