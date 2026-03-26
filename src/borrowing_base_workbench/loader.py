"""loader.py — Workbook-backed data loader for the detached Python engine.

Reads the borrowing base workbook using openpyxl (data_only mode) and returns
structured Python dataclasses.  No Excel COM, no PowerShell, no recalculation.

Classification of what is loaded here:
  RAW INPUT     — human-entered data; loaded as-is from the workbook cell
  POLICY CONFIG — facility-agreement thresholds; static per workbook version
  FORMULA PROXY — cell is an Excel formula, but the loader computes the same
                  value from already-loaded raw inputs (avoids cache staleness)
  MANUAL FLAG   — boolean field set by hand in Portfolio with no upstream source

NOT loaded here (computed by calculator.py in a later phase):
  - All Portfolio eligibility tests (cols E–N, P)
  - Collateral tier value (AF)
  - Assigned value (AL)
  - All concentration test columns (EI–HD)
  - All Availability outputs (L22, L23, L25, L26, L33–L35, L41–L42, L48)

Usage:
    from borrowing_base_workbench.loader import load_workbook_data
    data = load_workbook_data(Path("path/to/workbook.xlsm"))

    print(len(data.loans))        # LoanRecord list
    print(len(data.obligors))     # ObligorRecord list
    print(len(data.vaes))         # VaeRecord list
    print(data.policy)            # PolicyConfig
    print(data.availability_meta) # AvailabilityMeta
"""

from __future__ import annotations

import warnings
from dataclasses import dataclass, field
from datetime import date, datetime
from pathlib import Path
from typing import Any

from openpyxl import load_workbook as _openpyxl_load


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def _str(v: Any) -> str:
    return "" if v is None else str(v).strip()


def _float(v: Any) -> float | None:
    if v is None or (isinstance(v, str) and v.strip() in ("", "N/M", "N/A", "#N/A", "#VALUE!")):
        return None
    try:
        return float(v)
    except (ValueError, TypeError):
        return None


def _float_or(v: Any, default: float = 0.0) -> float:
    result = _float(v)
    return result if result is not None else default


def _date(v: Any) -> date | None:
    if v is None:
        return None
    if isinstance(v, datetime):
        return v.date()
    if isinstance(v, date):
        return v
    text = str(v).strip()
    for fmt in ("%Y-%m-%d", "%m/%d/%Y", "%m/%d/%y"):
        try:
            return datetime.strptime(text, fmt).date()
        except ValueError:
            continue
    return None


def _bool_yn(v: Any) -> bool:
    """Convert Yes/No/Y/N/TRUE/FALSE cell values to bool."""
    return str(v).strip().upper() in {"YES", "Y", "TRUE", "1", "-1"}


def _cov_lite(v: Any) -> bool:
    """Loan Tape col AC: -1 = cov-lite, 0 = has covenants."""
    if v is None:
        return False
    try:
        return int(float(v)) == -1
    except (ValueError, TypeError):
        return _bool_yn(v)


# ---------------------------------------------------------------------------
# Data classes
# ---------------------------------------------------------------------------

@dataclass
class LoanRecord:
    """One row from Loan Tape - Settled (rows 5–55).

    Source classification:
      RAW INPUT        : all fields unless noted
      FORMULA PROXY    : commitment_balance, olb, fmv,
                         current_ebitda_mm, current_senior_leverage,
                         current_interest_coverage, net_detachment, net_attachment
                         (computed from ObligorRecord by the loader)
    """
    # --- Identity (raw input) ---
    source_row: int                        # workbook row number (debug)
    ccm_id: str                            # col A
    fund_id: str                           # col C
    loan_id: int                           # col E
    obligor_name: str                      # col G  ← primary join key

    # --- Classification (raw input) ---
    security_display: str                  # col H  e.g. "1st Lien"
    loan_type: str                         # col J  "First Lien" | "FILO" | "Second Lien" | "DDTL" | "Revolver"
    security_type: str                     # col K  "Term Loan" | "Revolver" | "DDTL"
    denomination: str                      # col L

    # --- Balances (formula proxy — derived from SM Support via join) ---
    commitment_balance: float              # col M  = SM Support D
    olb: float                             # col N  = M for most rows
    unfunded: float                        # col O  raw input
    purchase_price: float                  # col Q  raw input; 1.0 = par

    # --- Valuation (formula proxy) ---
    fmv: float                             # col S  = SM Support E

    # --- Obligor info (raw input) ---
    country: str                           # col T
    investment_date: date | None           # col U
    maturity_date: date | None             # col V
    rate_type: str                         # col W  "Floating" | "Fixed"
    industry: str                          # col X  must match CL E81:E153 taxonomy
    payment_frequency: str                 # col Y  "M" | "Q"
    cov_lite: bool                         # col AC  -1 → True
    pik_pct: float                         # col AD
    spread: float                          # col AE
    sofr_floor: float                      # col AF
    is_ddtl: bool                          # col AG

    # --- Dates (raw input) ---
    inception_date: date | None            # col AN
    inclusion_date: date | None            # col AO  ⚠️ used in maturity test (BM-BL)/365

    # --- Initial financials (raw input — set at loan origination) ---
    initial_ebitda_mm: float               # col AV  stored in MILLIONS; loader multiplies ×1M
    initial_ebitda_nonadj_mm: float        # col AW  stored in MILLIONS
    initial_senior_leverage: float | None  # col AX
    initial_total_leverage: float | None   # col AY  ← used in eligibility Test 4
    initial_interest_coverage: float | None  # col BE

    # --- Current financials (formula proxy — from SM Support join) ---
    current_ebitda_mm: float               # col AZ  = SM Support G / 1,000,000
    current_senior_leverage: float | None  # col BC  = SM Support M
    current_interest_coverage: float | None  # col BF  = SM Support O
    net_detachment: float | None           # col BI  = SM Support M  (used for FILO bucketing)
    net_attachment: float | None           # col BJ  = SM Support N


@dataclass
class ObligorRecord:
    """One row from SM Support (rows 3–64).

    Source classification:
      RAW INPUT        : B, C, D, F, G, H, I, J, K, L, O, S, T, U, V, W, Y, Z
      FORMULA PROXY    : E (= D × AA), M (= (H+I-L)/G), N (= (H+I+J+K-L)/G), V (= S-T+U), AA (= AVG(Y,Z))
    """
    source_row: int                        # workbook row number (debug)
    obligor_name: str                      # col B  ← primary key

    # --- Identity ---
    security: str                          # col C

    # --- Balances (raw input) ---
    debt_balance: float                    # col D  BDC-specific investment
    sm_fair_value: float                   # col E  formula proxy = D × AA (mark)
    ltm_revenue: float                     # col F
    ltm_adj_ebitda: float                  # col G  in dollars (not millions)
    drawn_revolver: float                  # col H
    first_out: float                       # col I
    pari_passu: float                      # col J
    total_sm_balance: float                # col K
    cash_balance: float                    # col L

    # --- Leverage (formula proxy) ---
    net_detachment: float | None           # col M  = (H+I-L)/G   or None if EBITDA=0
    net_attachment: float | None           # col N  = (H+I+J+K-L)/G
    interest_coverage: float | None        # col O  raw input

    # --- PIK & funding (raw input) ---
    debt_balance_incl_pik: float           # col S
    pik_ltd: float                         # col T
    qtd_fundings: float                    # col U
    net_balance: float                     # col V  formula proxy = S-T+U

    # --- Valuation (raw input) ---
    mark_low: float | None                 # col Y
    mark_high: float | None                # col Z
    mark_avg: float | None                 # col AA  formula proxy = AVG(Y,Z)


@dataclass
class VaeRecord:
    """One row from VAE sheet (data rows start at row 5, headers at row 4).

    Source classification: all RAW INPUT
    """
    source_row: int
    borrower: str                          # col C  ← join key to LoanRecord.obligor_name
    event_type: str                        # col D
    specific_test: str                     # col E
    material_modification: str            # col F  raw text (n/a, Yes, No)
    vae_date: date | None                  # col G
    ebitda_at_vae: float | None            # col H
    agent_adj_haircut: float | None        # col I
    permitted_ebitda: float | None         # col J  = H - I
    senior_debt: float | None              # col K
    total_debt: float | None               # col L
    net_senior_leverage: float | None      # col M
    net_total_leverage: float | None       # col N
    interest_coverage_at_vae: float | None # col O  (may be "N/A")
    vae_agent_assigned_value: float | None # col P  ← operative value for engine


@dataclass
class ManualPortfolioFlags:
    """Manual flags set per loan row in the Portfolio sheet.

    These have no upstream formula chain.  For existing loans, they are
    loaded directly from Portfolio.  For scenario loans, all default False.
    """
    obligor_name: str                      # from Portfolio col B (join key)
    limited_industry: bool                 # col BV  Test 7
    non_sponsor: bool                      # col BW  Test 9a
    div_recap: bool                        # col BX  Test 9b
    is_dip: bool                           # col CA  Test 2 (PIK/DIP)
    agent_addback_discretion: bool         # col CQ  affects permitted EBITDA (CY)
    agent_post_inclusion_haircut: bool     # col CW  affects permitted EBITDA (CY)
    agent_addback_haircut_pct: float | None  # col CX  quantified haircut if CW = Y


@dataclass
class CollateralTier:
    """One row from AGENT assigned value grid (rows 49–52)."""
    max_first_lien_lev: float              # col B threshold (senior net leverage)
    max_filo_2l_lev: float                 # col C threshold (total net leverage)
    applicable_pct: float                  # col D value (0, 0.85, 0.925, 1.0)


@dataclass
class PolicyConfig:
    """All thresholds and rates loaded from AGENT sheet.

    Source classification: all POLICY CONFIG (static per facility agreement).
    """
    # --- Eligibility thresholds ---
    max_total_leverage_at_inclusion: float     # AGENT D4
    min_ebitda_second_lien_mm: float           # AGENT B4  (in millions, e.g. 10 → $10MM)
    min_acquisition_price: float               # AGENT E4
    max_original_maturity_years: float         # AGENT F4

    # --- Advance rates ---
    rate_first_lien_large: float               # AGENT D23  EBITDA > $20MM
    rate_first_lien_mid: float                 # AGENT D24  EBITDA $10–20MM
    rate_first_lien_small: float               # AGENT D25  EBITDA < $10MM
    rate_filo_low_lev: float                   # AGENT D27  net det < 0.75x
    rate_filo_mid_lev: float                   # AGENT D28  net det 0.75–1.25x
    rate_filo_high_lev: float                  # AGENT D29  net det > 1.25x
    rate_second_lien: float                    # AGENT D31

    # --- Portfolio advance rate caps ---
    cap_high_diversity: float                  # AGENT D41  >= 20 obligors
    cap_mid_diversity: float                   # AGENT D42  12–20 obligors
    cap_low_diversity: float                   # AGENT D43  < 12 obligors
    threshold_high_diversity: int              # AGENT B41  20
    threshold_mid_diversity: int               # AGENT B42  12

    # --- Collateral value tiers ---
    collateral_tiers: list[CollateralTier]     # AGENT rows 49–52

    # --- Concentration limits (13 tests) ---
    conc_max_2l_filo_high_lev: float           # AGENT D57
    conc_max_second_lien: float                # AGENT D58
    conc_max_non_first_lien: float             # AGENT D59
    conc_max_small_ebitda: float               # AGENT D60
    conc_max_obligor: float                    # AGENT D61
    conc_max_largest_industry: float           # AGENT D62
    conc_max_second_industry: float            # AGENT D63
    conc_max_other_industries: float           # AGENT D64
    conc_fixed_rate: float                     # AGENT D65
    conc_limited_industry: float               # AGENT D66
    conc_max_ddtl_revolver: float              # AGENT D67
    conc_max_non_sponsor: float                # AGENT D68
    conc_max_div_recap: float                  # AGENT D69

    # --- Limited industries ---
    limited_industries: list[str]              # AGENT A73:A76  (lowercase)

    # --- Industry taxonomy (from Concentration Limits E81:E153) ---
    industry_taxonomy: list[str]               # 73 S&P industry names


@dataclass
class AvailabilityMeta:
    """Minimal metadata loaded from Availability sheet.

    Only the three raw/constant cells are loaded here.
    All formula-driven outputs (L26, L33–L35, L41–L42, L48) are
    computed by calculator.py — never read from the cached workbook.
    """
    facility_amount: float                 # L16  raw input / policy constant
    available_collections: float          # L18  cash in Principal Collection Account
    current_advances: float               # L51  hardcoded constant in workbook
    measurement_date: date | None         # F8   reporting date


@dataclass
class WorkbookData:
    """Top-level container for all workbook-backed data.

    Returned by load_workbook_data(). This is the input to engine.py.
    """
    workbook_path: str
    loans: list[LoanRecord]
    obligors: list[ObligorRecord]
    vaes: list[VaeRecord]
    portfolio_flags: dict[str, ManualPortfolioFlags]   # keyed by obligor_name
    policy: PolicyConfig
    availability_meta: AvailabilityMeta

    def obligor_by_name(self, name: str) -> ObligorRecord | None:
        for o in self.obligors:
            if o.obligor_name == name:
                return o
        return None

    def vaes_for(self, obligor_name: str) -> list[VaeRecord]:
        return [v for v in self.vaes if v.borrower == obligor_name]

    def flags_for(self, obligor_name: str) -> ManualPortfolioFlags:
        return self.portfolio_flags.get(
            obligor_name,
            ManualPortfolioFlags(
                obligor_name=obligor_name,
                limited_industry=False, non_sponsor=False, div_recap=False,
                is_dip=False, agent_addback_discretion=False,
                agent_post_inclusion_haircut=False, agent_addback_haircut_pct=None,
            ),
        )

    def summary(self) -> dict:
        return {
            "workbook_path": self.workbook_path,
            "loan_count": len(self.loans),
            "obligor_count": len(self.obligors),
            "vae_count": len(self.vaes),
            "portfolio_flags_count": len(self.portfolio_flags),
            "facility_amount": self.availability_meta.facility_amount,
            "current_advances": self.availability_meta.current_advances,
            "measurement_date": str(self.availability_meta.measurement_date),
            "industry_taxonomy_count": len(self.policy.industry_taxonomy),
            "limited_industries": self.policy.limited_industries,
        }


# ---------------------------------------------------------------------------
# Sheet loaders
# ---------------------------------------------------------------------------

def _load_loan_tape(ws, obligor_lookup: dict[str, ObligorRecord]) -> list[LoanRecord]:
    """Load Loan Tape - Settled rows 5–55.

    Formula-proxy columns (M, S, AZ, BC, BF, BI, BJ) are computed from
    the obligor_lookup (SM Support) to avoid relying on cached formula values.
    """
    loans: list[LoanRecord] = []

    for row in range(5, 56):
        obligor_name = _str(ws[f"G{row}"].value)
        if not obligor_name:
            continue

        obs = obligor_lookup.get(obligor_name)

        # formula-proxy values: pull from SM Support if available, else cached cell
        commitment_balance = (obs.debt_balance if obs else None) or _float_or(ws[f"M{row}"].value)
        fmv               = (obs.sm_fair_value if obs else None) or _float_or(ws[f"S{row}"].value)
        current_ebitda_mm = (obs.ltm_adj_ebitda / 1_000_000 if obs else None) if obs else (
            _float_or(ws[f"AZ{row}"].value)  # cached: already in millions
        )
        current_senior_lev = (obs.net_detachment if obs else None) if obs else _float(ws[f"BC{row}"].value)
        current_interest_cov = (obs.interest_coverage if obs else None) if obs else _float(ws[f"BF{row}"].value)
        net_detachment = (obs.net_detachment if obs else None) if obs else _float(ws[f"BI{row}"].value)
        net_attachment = (obs.net_attachment if obs else None) if obs else _float(ws[f"BJ{row}"].value)

        initial_ebitda_raw = _float(ws[f"AV{row}"].value)
        initial_ebitda_mm  = (initial_ebitda_raw * 1_000_000) if initial_ebitda_raw is not None else 0.0
        initial_ebitda_nonadj_raw = _float(ws[f"AW{row}"].value)
        initial_ebitda_nonadj_mm  = (initial_ebitda_nonadj_raw * 1_000_000) if initial_ebitda_nonadj_raw is not None else 0.0

        loans.append(LoanRecord(
            source_row=row,
            ccm_id=_str(ws[f"A{row}"].value),
            fund_id=_str(ws[f"C{row}"].value),
            loan_id=int(_float_or(ws[f"E{row}"].value)),
            obligor_name=obligor_name,
            security_display=_str(ws[f"H{row}"].value),
            loan_type=_str(ws[f"J{row}"].value),
            security_type=_str(ws[f"K{row}"].value),
            denomination=_str(ws[f"L{row}"].value),
            commitment_balance=commitment_balance,
            olb=commitment_balance,  # N = M for settled loans
            unfunded=_float_or(ws[f"O{row}"].value),
            purchase_price=_float_or(ws[f"Q{row}"].value, default=1.0),
            fmv=fmv,
            country=_str(ws[f"T{row}"].value),
            investment_date=_date(ws[f"U{row}"].value),
            maturity_date=_date(ws[f"V{row}"].value),
            rate_type=_str(ws[f"W{row}"].value),
            industry=_str(ws[f"X{row}"].value),
            payment_frequency=_str(ws[f"Y{row}"].value),
            cov_lite=_cov_lite(ws[f"AC{row}"].value),
            pik_pct=_float_or(ws[f"AD{row}"].value),
            spread=_float_or(ws[f"AE{row}"].value),
            sofr_floor=_float_or(ws[f"AF{row}"].value),
            is_ddtl=_bool_yn(ws[f"AG{row}"].value),
            inception_date=_date(ws[f"AN{row}"].value),
            inclusion_date=_date(ws[f"AO{row}"].value),
            initial_ebitda_mm=initial_ebitda_mm,
            initial_ebitda_nonadj_mm=initial_ebitda_nonadj_mm,
            initial_senior_leverage=_float(ws[f"AX{row}"].value),
            initial_total_leverage=_float(ws[f"AY{row}"].value),
            initial_interest_coverage=_float(ws[f"BE{row}"].value),
            current_ebitda_mm=current_ebitda_mm or 0.0,
            current_senior_leverage=current_senior_lev,
            current_interest_coverage=current_interest_cov,
            net_detachment=net_detachment,
            net_attachment=net_attachment,
        ))

    return loans


def _load_sm_support(ws) -> list[ObligorRecord]:
    """Load SM Support rows 3–64.

    M, N computed from raw inputs (H,I,J,K,L,G).
    E computed from D × AA.
    V computed from S - T + U.
    """
    obligors: list[ObligorRecord] = []

    for row in range(3, 65):
        name = _str(ws[f"B{row}"].value)
        if not name:
            continue

        g = _float_or(ws[f"G{row}"].value)  # EBITDA
        h = _float_or(ws[f"H{row}"].value)  # revolver
        i = _float_or(ws[f"I{row}"].value)  # first out
        j = _float_or(ws[f"J{row}"].value)  # pari
        k = _float_or(ws[f"K{row}"].value)  # total SM
        l = _float_or(ws[f"L{row}"].value)  # cash

        net_det = ((h + i - l) / g) if g != 0 else None
        net_att = ((h + i + j + k - l) / g) if g != 0 else None

        d  = _float_or(ws[f"D{row}"].value)  # debt balance
        y_ = _float(ws[f"Y{row}"].value)     # mark low
        z_ = _float(ws[f"Z{row}"].value)     # mark high
        aa = ((y_ + z_) / 2) if (y_ is not None and z_ is not None) else None
        fmv = (d * aa) if (aa is not None and d) else d

        s = _float_or(ws[f"S{row}"].value)
        t = _float_or(ws[f"T{row}"].value)
        u = _float_or(ws[f"U{row}"].value)

        obligors.append(ObligorRecord(
            source_row=row,
            obligor_name=name,
            security=_str(ws[f"C{row}"].value),
            debt_balance=d,
            sm_fair_value=fmv if fmv is not None else d,
            ltm_revenue=_float_or(ws[f"F{row}"].value),
            ltm_adj_ebitda=g,
            drawn_revolver=h,
            first_out=i,
            pari_passu=j,
            total_sm_balance=k,
            cash_balance=l,
            net_detachment=net_det,
            net_attachment=net_att,
            interest_coverage=_float(ws[f"O{row}"].value),
            debt_balance_incl_pik=s,
            pik_ltd=t,
            qtd_fundings=u,
            net_balance=s - t + u,
            mark_low=y_,
            mark_high=z_,
            mark_avg=aa,
        ))

    return obligors


def _load_vae(ws) -> list[VaeRecord]:
    """Load VAE sheet — header row 4, data rows 5+."""
    vaes: list[VaeRecord] = []

    for row in range(5, 200):
        borrower = _str(ws[f"C{row}"].value)
        if not borrower:
            break

        vaes.append(VaeRecord(
            source_row=row,
            borrower=borrower,
            event_type=_str(ws[f"D{row}"].value),
            specific_test=_str(ws[f"E{row}"].value),
            material_modification=_str(ws[f"F{row}"].value),
            vae_date=_date(ws[f"G{row}"].value),
            ebitda_at_vae=_float(ws[f"H{row}"].value),
            agent_adj_haircut=_float(ws[f"I{row}"].value),
            permitted_ebitda=_float(ws[f"J{row}"].value),
            senior_debt=_float(ws[f"K{row}"].value),
            total_debt=_float(ws[f"L{row}"].value),
            net_senior_leverage=_float(ws[f"M{row}"].value),
            net_total_leverage=_float(ws[f"N{row}"].value),
            interest_coverage_at_vae=_float(ws[f"O{row}"].value),
            vae_agent_assigned_value=_float(ws[f"P{row}"].value),
        ))

    return vaes


def _load_policy(agent_ws, cl_ws) -> PolicyConfig:
    """Load AGENT sheet thresholds + Concentration Limits industry taxonomy."""
    ag = agent_ws

    tiers: list[CollateralTier] = []
    for row in range(49, 53):
        b = _float(ag[f"B{row}"].value)
        c = _float(ag[f"C{row}"].value)
        d = _float(ag[f"D{row}"].value)
        if b is not None and c is not None and d is not None:
            tiers.append(CollateralTier(
                max_first_lien_lev=b,
                max_filo_2l_lev=c,
                applicable_pct=d,
            ))
    # Row 52 stores the 0% threshold as a text row ("portion of 0%...").
    # Always append the sentinel: leverage >= 5.0x -> 0% collateral value.
    if not any(t.applicable_pct == 0.0 for t in tiers):
        tiers.append(CollateralTier(max_first_lien_lev=float("inf"), max_filo_2l_lev=float("inf"), applicable_pct=0.0))

    # Industry taxonomy from Concentration Limits E81:E153
    taxonomy: list[str] = []
    for row in range(81, 154):
        val = _str(cl_ws[f"E{row}"].value)
        if val and "Max " not in val and val != "Industry":
            taxonomy.append(val)

    limited_industries = [
        _str(ag[f"A{r}"].value).lower()
        for r in range(73, 77)
        if _str(ag[f"A{r}"].value)
    ]

    return PolicyConfig(
        max_total_leverage_at_inclusion=_float_or(ag["D4"].value, 6.5),
        min_ebitda_second_lien_mm=_float_or(ag["B4"].value, 10.0),
        min_acquisition_price=_float_or(ag["E4"].value, 0.9),
        max_original_maturity_years=_float_or(ag["F4"].value, 7.0),
        rate_first_lien_large=_float_or(ag["D23"].value, 0.675),
        rate_first_lien_mid=_float_or(ag["D24"].value, 0.65),
        rate_first_lien_small=_float_or(ag["D25"].value, 0.625),
        rate_filo_low_lev=_float_or(ag["D27"].value, 0.625),
        rate_filo_mid_lev=_float_or(ag["D28"].value, 0.575),
        rate_filo_high_lev=_float_or(ag["D29"].value, 0.45),
        rate_second_lien=_float_or(ag["D31"].value, 0.45),
        cap_high_diversity=_float_or(ag["D41"].value, 0.65),
        cap_mid_diversity=_float_or(ag["D42"].value, 0.625),
        cap_low_diversity=_float_or(ag["D43"].value, 0.55),
        threshold_high_diversity=int(_float_or(ag["B41"].value, 20)),
        threshold_mid_diversity=int(_float_or(ag["B42"].value, 12)),
        collateral_tiers=tiers,
        conc_max_2l_filo_high_lev=_float_or(ag["D57"].value, 0.20),
        conc_max_second_lien=_float_or(ag["D58"].value, 0.10),
        conc_max_non_first_lien=_float_or(ag["D59"].value, 0.30),
        conc_max_small_ebitda=_float_or(ag["D60"].value, 0.15),
        conc_max_obligor=_float_or(ag["D61"].value, 0.075),
        conc_max_largest_industry=_float_or(ag["D62"].value, 0.20),
        conc_max_second_industry=_float_or(ag["D63"].value, 0.15),
        conc_max_other_industries=_float_or(ag["D64"].value, 0.10),
        conc_fixed_rate=_float_or(ag["D65"].value, 0.10),
        conc_limited_industry=_float_or(ag["D66"].value, 0.10),
        conc_max_ddtl_revolver=_float_or(ag["D67"].value, 0.15),
        conc_max_non_sponsor=_float_or(ag["D68"].value, 0.15),
        conc_max_div_recap=_float_or(ag["D69"].value, 0.10),
        limited_industries=limited_industries,
        industry_taxonomy=taxonomy,
    )


def _load_portfolio_flags(ws, obligor_ws) -> dict[str, ManualPortfolioFlags]:
    """Load manual boolean flags from Portfolio sheet (cols BV, BW, BX, CA, CQ, CW, CX).

    Keyed by obligor_name (col B).
    """
    flags: dict[str, ManualPortfolioFlags] = {}

    for row in range(9, 99):
        name = _str(ws[f"B{row}"].value)
        if not name:
            continue
        flags[name] = ManualPortfolioFlags(
            obligor_name=name,
            limited_industry=_bool_yn(ws[f"BV{row}"].value),
            non_sponsor=_bool_yn(ws[f"BW{row}"].value),
            div_recap=_bool_yn(ws[f"BX{row}"].value),
            is_dip=_bool_yn(ws[f"CA{row}"].value),
            agent_addback_discretion=_str(ws[f"CQ{row}"].value).upper() == "Y",
            agent_post_inclusion_haircut=_str(ws[f"CW{row}"].value).upper() == "Y",
            agent_addback_haircut_pct=_float(ws[f"CX{row}"].value),
        )

    return flags


def _load_availability_meta(ws) -> AvailabilityMeta:
    """Load the three raw/constant cells from Availability sheet."""
    return AvailabilityMeta(
        facility_amount=_float_or(ws["L16"].value, 200_000_000.0),
        available_collections=_float_or(ws["L18"].value, 0.0),
        current_advances=_float_or(ws["L51"].value, 0.0),
        measurement_date=_date(ws["F8"].value),
    )


# ---------------------------------------------------------------------------
# Public entry point
# ---------------------------------------------------------------------------

def load_workbook_data(workbook_path: str | Path) -> WorkbookData:
    """Open the workbook and return all raw/policy data as structured Python objects.

    Uses openpyxl in data_only mode (reads cached cell values, no recalculation).
    Workbook must have been saved after a full calculation for cached values to
    reflect current formula results.

    Args:
        workbook_path: path to the .xlsx or .xlsm workbook file

    Returns:
        WorkbookData containing loans, obligors, vaes, policy, flags, meta
    """
    workbook_path = Path(workbook_path)

    with warnings.catch_warnings():
        warnings.filterwarnings("ignore", category=UserWarning)
        wb = _openpyxl_load(workbook_path, data_only=True, keep_vba=False, read_only=True)

    # Load SM Support first — needed to resolve formula-proxy cols in Loan Tape
    obligors = _load_sm_support(wb["SM Support"])
    obligor_lookup = {o.obligor_name: o for o in obligors}

    loans = _load_loan_tape(wb["Loan Tape - Settled"], obligor_lookup)
    vaes  = _load_vae(wb["VAE"])
    flags = _load_portfolio_flags(wb["Portfolio"], wb["SM Support"])
    policy = _load_policy(wb["AGENT"], wb["Concentration Limits"])
    meta   = _load_availability_meta(wb["Availability"])

    wb.close()

    return WorkbookData(
        workbook_path=str(workbook_path),
        loans=loans,
        obligors=obligors,
        vaes=vaes,
        portfolio_flags=flags,
        policy=policy,
        availability_meta=meta,
    )


# ---------------------------------------------------------------------------
# CLI test / debug path
# ---------------------------------------------------------------------------

if __name__ == "__main__":
    import sys
    import json as _json

    path = Path(sys.argv[1]) if len(sys.argv) > 1 else Path(
        r"C:\Users\henry.yan\Downloads\2025-02-11_BDC Borrowing_Base_v8.xlsm"
    )

    print(f"Loading: {path}")
    data = load_workbook_data(path)

    summary = data.summary()
    print("\n=== Summary ===")
    for k, v in summary.items():
        print(f"  {k}: {v}")

    print(f"\n=== Loans ({len(data.loans)}) ===")
    for loan in data.loans[:5]:
        print(
            f"  row={loan.source_row} | {loan.obligor_name!r:45s} | "
            f"{loan.loan_type:12s} | olb=${loan.olb:>15,.0f} | "
            f"industry={loan.industry!r}"
        )
    if len(data.loans) > 5:
        print(f"  ... ({len(data.loans) - 5} more)")

    print(f"\n=== Obligors ({len(data.obligors)}) ===")
    for o in data.obligors[:5]:
        nd = f"{o.net_detachment:.3f}" if o.net_detachment is not None else "None"
        print(
            f"  row={o.source_row} | {o.obligor_name!r:45s} | "
            f"ebitda=${o.ltm_adj_ebitda:>12,.0f} | net_det={nd}"
        )
    if len(data.obligors) > 5:
        print(f"  ... ({len(data.obligors) - 5} more)")

    print(f"\n=== VAEs ({len(data.vaes)}) ===")
    for v in data.vaes[:5]:
        print(
            f"  row={v.source_row} | {v.borrower!r:40s} | "
            f"type={v.event_type!r:20s} | date={v.vae_date} | "
            f"assigned_value={v.vae_agent_assigned_value}"
        )
    if len(data.vaes) > 5:
        print(f"  ... ({len(data.vaes) - 5} more)")

    print(f"\n=== Policy ===")
    print(f"  max_total_lev_at_inclusion : {data.policy.max_total_leverage_at_inclusion}")
    print(f"  min_2L_ebitda ($MM)        : {data.policy.min_ebitda_second_lien_mm}")
    print(f"  min_acquisition_price      : {data.policy.min_acquisition_price}")
    print(f"  max_maturity_years         : {data.policy.max_original_maturity_years}")
    print(f"  advance rates              : FL_lg={data.policy.rate_first_lien_large} "
          f"FL_md={data.policy.rate_first_lien_mid} FL_sm={data.policy.rate_first_lien_small} "
          f"2L={data.policy.rate_second_lien}")
    print(f"  collateral_tiers           : {data.policy.collateral_tiers}")
    print(f"  limited_industries         : {data.policy.limited_industries}")
    print(f"  industry_taxonomy          : {len(data.policy.industry_taxonomy)} entries")

    print(f"\n=== Availability Meta ===")
    print(f"  facility_amount     : ${data.availability_meta.facility_amount:,.0f}")
    print(f"  available_collections: ${data.availability_meta.available_collections:,.0f}")
    print(f"  current_advances    : ${data.availability_meta.current_advances:,.0f}")
    print(f"  measurement_date    : {data.availability_meta.measurement_date}")

    print(f"\n=== Portfolio Flags (non-default sample) ===")
    non_default = {
        k: v for k, v in data.portfolio_flags.items()
        if v.non_sponsor or v.limited_industry or v.div_recap or v.is_dip
    }
    if non_default:
        for name, f in list(non_default.items())[:10]:
            print(
                f"  {name!r:45s} | non_sponsor={f.non_sponsor} "
                f"limited_ind={f.limited_industry} div_recap={f.div_recap} dip={f.is_dip}"
            )
    else:
        print("  (all flags at default False for current portfolio)")

    print("\nDone.")
