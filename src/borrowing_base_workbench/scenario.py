"""scenario.py — Scenario record construction for pro forma runs.

Converts the 23-key string dict collected from the frontend form into typed
LoanRecord + ObligorRecord + ManualPortfolioFlags, then injects those records
into a WorkbookData copy for a clean engine rerun.

Field mapping notes
-------------------
Form field          → LoanRecord / ObligorRecord field
--------------------  --------------------------------
company_name        → obligor_name (both records)
security_type       → loan_type  ("First Lien" | "FILO"; the form combobox
                       uses the same values as the CLA loan-type taxonomy)
ltm_revenue         → ObligorRecord.ltm_revenue  (dollars)
ltm_adj_ebitda      → ObligorRecord.ltm_adj_ebitda  (dollars)
                       LoanRecord.current_ebitda_mm = /1_000_000  (millions)
drawn_revolver      → ObligorRecord.drawn_revolver  (dollars)
first_out_balance   → ObligorRecord.first_out  (dollars)
pari_passu          → ObligorRecord.pari_passu  (dollars)
bdc_balance         → ObligorRecord.debt_balance
                       LoanRecord.olb, LoanRecord.commitment_balance
total_sm_balance    → ObligorRecord.total_sm_balance  (dollars)
cash_balance        → ObligorRecord.cash_balance  (dollars)
interest_coverage   → ObligorRecord.interest_coverage  (raw multiple)
loan_denomination   → LoanRecord.denomination
purchase_price      → LoanRecord.purchase_price  (fraction, e.g. 1.0 = par)
                       FMV = purchase_price × bdc_balance
country             → LoanRecord.country
investment_date     → LoanRecord.investment_date, .inclusion_date, .inception_date
maturity_date       → LoanRecord.maturity_date
rate_type           → LoanRecord.rate_type  ("Floating" | "Fixed")
industry_class…     → LoanRecord.industry
payment_frequency   → LoanRecord.payment_frequency  ("M" | "Q")
pik_pct             → LoanRecord.pik_pct  (fraction)
spread              → LoanRecord.spread  (fraction)
sofr_floor          → LoanRecord.sofr_floor  (fraction)
attach_point        → used only as a leverage fallback when EBITDA = 0

Leverage derivation (mirrors SM Support formula-proxy columns M and N):
    net_detachment = (drawn_revolver + first_out - cash_balance) / ltm_adj_ebitda
    net_attachment = (drawn_revolver + first_out + pari_passu + total_sm - cash) / ltm_adj_ebitda

All scenario ManualPortfolioFlags default to False (no sponsor, no div-recap,
no limited industry, no DIP).
"""

from __future__ import annotations

import dataclasses
from datetime import date, datetime
from typing import Any

from borrowing_base_workbench.loader import (
    LoanRecord,
    ManualPortfolioFlags,
    ObligorRecord,
    WorkbookData,
)

# Sentinel row number used to tag synthetic records.
# Must not collide with any real workbook row (1–200 used).
_SYNTHETIC_ROW = 9999


# ---------------------------------------------------------------------------
# Parsing helpers (all scenario values arrive as plain strings)
# ---------------------------------------------------------------------------

def _float(v: Any, default: float = 0.0) -> float:
    """Parse a string/number to float; return default on failure."""
    if v is None or v == "":
        return default
    try:
        cleaned = (
            str(v).strip()
            .replace("$", "").replace(",", "")
            .replace("%", "").replace("x", "").replace("X", "")
        )
        return float(cleaned) if cleaned else default
    except (ValueError, TypeError):
        return default


def _float_opt(v: Any) -> float | None:
    """Parse to float or None."""
    if v is None or v == "":
        return None
    try:
        cleaned = (
            str(v).strip()
            .replace("$", "").replace(",", "")
            .replace("%", "").replace("x", "").replace("X", "")
        )
        return float(cleaned) if cleaned else None
    except (ValueError, TypeError):
        return None


def _date(v: Any) -> date | None:
    """Parse a date string."""
    if not v:
        return None
    if isinstance(v, date):
        return v
    text = str(v).strip()
    for fmt in ("%Y-%m-%d", "%m/%d/%Y", "%m/%d/%y"):
        try:
            return datetime.strptime(text, fmt).date()
        except ValueError:
            continue
    return None


# ---------------------------------------------------------------------------
# Public API
# ---------------------------------------------------------------------------

def build_scenario_records(
    scenario: dict[str, str],
) -> tuple[LoanRecord, ObligorRecord, ManualPortfolioFlags]:
    """Convert a scenario form dict → (LoanRecord, ObligorRecord, ManualPortfolioFlags).

    All dict values are strings as supplied by the Tkinter form (currency fields
    are plain float strings, percent fields are fractions, multiples are floats).

    Returns three records ready to be injected into WorkbookData.
    """
    name = scenario.get("company_name", "").strip() or "Scenario Company"

    # --- SM Support / ObligorRecord values ----------------------------------
    ltm_revenue      = _float(scenario.get("ltm_revenue"))
    ltm_adj_ebitda   = _float(scenario.get("ltm_adj_ebitda"))   # dollars
    drawn_revolver   = _float(scenario.get("drawn_revolver"))
    first_out        = _float(scenario.get("first_out_balance"))
    pari_passu       = _float(scenario.get("pari_passu"))
    bdc_balance      = _float(scenario.get("bdc_balance"))
    total_sm_balance = _float(scenario.get("total_sm_balance"))
    cash_balance     = _float(scenario.get("cash_balance"))
    interest_cov     = _float_opt(scenario.get("interest_coverage"))

    # Leverage from financial inputs — mirrors SM Support formula-proxy cols M/N
    if ltm_adj_ebitda > 0:
        net_det: float | None = (drawn_revolver + first_out - cash_balance) / ltm_adj_ebitda
        net_att: float | None = (
            drawn_revolver + first_out + pari_passu + total_sm_balance - cash_balance
        ) / ltm_adj_ebitda
    else:
        # Fallback: use attach_point as a direct leverage proxy
        fallback = _float_opt(scenario.get("attach_point"))
        net_det = fallback
        net_att = fallback

    # Mark: purchase_price fraction used as mark (new investments are typically at par)
    purchase_price  = _float(scenario.get("purchase_price"), default=1.0)
    sm_fair_value   = bdc_balance * purchase_price

    obligor = ObligorRecord(
        source_row=_SYNTHETIC_ROW,
        obligor_name=name,
        security=scenario.get("security_type", ""),
        debt_balance=bdc_balance,
        sm_fair_value=sm_fair_value,
        ltm_revenue=ltm_revenue,
        ltm_adj_ebitda=ltm_adj_ebitda,
        drawn_revolver=drawn_revolver,
        first_out=first_out,
        pari_passu=pari_passu,
        total_sm_balance=total_sm_balance,
        cash_balance=cash_balance,
        net_detachment=net_det,
        net_attachment=net_att,
        interest_coverage=interest_cov,
        # Remaining SM Support fields — not available from scenario form
        debt_balance_incl_pik=bdc_balance,
        pik_ltd=0.0,
        qtd_fundings=0.0,
        net_balance=bdc_balance,
        mark_low=purchase_price,
        mark_high=purchase_price,
        mark_avg=purchase_price,
    )

    # --- Loan Tape / LoanRecord values --------------------------------------
    # The form "security_type" combobox surfaces "First Lien" / "FILO" which
    # map directly to LoanRecord.loan_type in the CLA taxonomy.
    loan_type    = scenario.get("security_type", "First Lien")
    denomination = scenario.get("loan_denomination", "USD")
    country      = scenario.get("country", "United States")
    invest_dt    = _date(scenario.get("investment_date"))
    maturity_dt  = _date(scenario.get("maturity_date"))
    rate_type    = scenario.get("rate_type", "Floating")
    industry     = scenario.get("industry_classification", "")
    pay_freq     = scenario.get("payment_frequency", "M")
    pik_pct      = _float(scenario.get("pik_pct"))
    spread       = _float(scenario.get("spread"))
    sofr_floor   = _float(scenario.get("sofr_floor"))

    # initial_ebitda_mm: stored in dollars (loader multiplied workbook millions × 1M)
    # Use current EBITDA as initial — reasonable for a new origination
    initial_ebitda_dollars = ltm_adj_ebitda

    loan = LoanRecord(
        source_row=_SYNTHETIC_ROW,
        ccm_id="SCENARIO",
        fund_id="SCENARIO",
        loan_id=0,
        obligor_name=name,
        security_display=loan_type,
        loan_type=loan_type,
        security_type="Term Loan",      # synthetic loans are term loans
        denomination=denomination,
        commitment_balance=bdc_balance,
        olb=bdc_balance,
        unfunded=0.0,
        purchase_price=purchase_price,
        fmv=sm_fair_value,
        country=country,
        investment_date=invest_dt,
        maturity_date=maturity_dt,
        rate_type=rate_type,
        industry=industry,
        payment_frequency=pay_freq,
        cov_lite=False,
        pik_pct=pik_pct,
        spread=spread,
        sofr_floor=sofr_floor,
        is_ddtl=(loan_type == "DDTL"),
        inception_date=invest_dt,
        # New loans: inclusion_date == investment_date (no workbook history)
        inclusion_date=invest_dt,
        initial_ebitda_mm=initial_ebitda_dollars,
        initial_ebitda_nonadj_mm=initial_ebitda_dollars,
        initial_senior_leverage=net_det,
        initial_total_leverage=net_att,  # total leverage proxy for eligibility Test H
        initial_interest_coverage=interest_cov,
        # current_ebitda_mm is in MILLIONS (mirrors SM Support G / 1,000,000)
        current_ebitda_mm=(ltm_adj_ebitda / 1_000_000) if ltm_adj_ebitda else 0.0,
        current_senior_leverage=net_det,
        current_interest_coverage=interest_cov,
        net_detachment=net_det,
        net_attachment=net_att,
    )

    flags = ManualPortfolioFlags(
        obligor_name=name,
        limited_industry=False,
        non_sponsor=False,
        div_recap=False,
        is_dip=False,
        agent_addback_discretion=False,
        agent_post_inclusion_haircut=False,
        agent_addback_haircut_pct=None,
    )

    return loan, obligor, flags


def inject_scenario(
    data: WorkbookData,
    loan: LoanRecord,
    obligor: ObligorRecord,
    flags: ManualPortfolioFlags,
) -> WorkbookData:
    """Return a new WorkbookData with the synthetic records appended.

    No mutation: returns a shallow copy with the new records appended to lists
    and flags dict.  The original WorkbookData is unchanged.
    """
    new_flags = {**data.portfolio_flags, obligor.obligor_name: flags}
    return dataclasses.replace(
        data,
        loans=[*data.loans, loan],
        obligors=[*data.obligors, obligor],
        portfolio_flags=new_flags,
    )


# ---------------------------------------------------------------------------
# CLI smoke test
# ---------------------------------------------------------------------------

if __name__ == "__main__":
    import sys
    from pathlib import Path

    from borrowing_base_workbench.calculator import calculate_portfolio
    from borrowing_base_workbench.loader import load_workbook_data

    wb_path = Path(r"C:\Users\henry.yan\Downloads\2025-02-11_BDC Borrowing_Base_v8.xlsm")
    if len(sys.argv) > 1:
        wb_path = Path(sys.argv[1])

    print(f"Loading baseline: {wb_path}")
    data = load_workbook_data(wb_path)
    before = calculate_portfolio(data)

    # Minimal example scenario — First Lien, $25MM EBITDA, $20MM BDC balance
    example_scenario = {
        "company_name":          "Acme Corp (Pro Forma)",
        "security_type":         "First Lien",
        "ltm_revenue":           "100000000",    # $100MM
        "ltm_adj_ebitda":        "25000000",     # $25MM
        "drawn_revolver":        "0",
        "first_out_balance":     "0",
        "pari_passu":            "0",
        "bdc_balance":           "20000000",     # $20MM
        "total_sm_balance":      "20000000",
        "cash_balance":          "0",
        "interest_coverage":     "2.5",
        "loan_denomination":     "USD",
        "purchase_price":        "1.0",          # par
        "country":               "United States",
        "investment_date":       "2025-01-15",
        "maturity_date":         "2031-01-15",   # 6-year term
        "rate_type":             "Floating",
        "industry_classification": "Software",
        "payment_frequency":     "M",
        "pik_pct":               "0",
        "spread":                "0.055",        # 5.5%
        "sofr_floor":            "0.005",        # 0.5%
        "attach_point":          "0.0",          # not needed; EBITDA > 0
    }

    loan, obligor, flags = build_scenario_records(example_scenario)
    after_data = inject_scenario(data, loan, obligor, flags)
    after = calculate_portfolio(after_data)

    SEP = "=" * 70
    print(f"\n{SEP}")
    print("  PHASE 5 — PRO FORMA SCENARIO RESULT")
    print(SEP)
    print(f"  Scenario loan:   {loan.obligor_name}  ({loan.loan_type})")
    print(f"  OLB:             ${loan.olb:>12,.0f}")
    print(f"  Current EBITDA:  ${loan.current_ebitda_mm:>12.1f}MM")
    print(f"  Net detachment:  {loan.net_detachment:.2f}x" if loan.net_detachment is not None else "  Net detachment:  N/A")

    # Eligibility of scenario loan
    for asset in after.assets:
        if asset.obligor_name == loan.obligor_name:
            e = asset.eligibility
            print(f"\n  Eligibility:     {'ELIGIBLE' if e.eligible else 'INELIGIBLE'}")
            if e.failed_tests:
                for t in e.failed_tests:
                    print(f"    FAIL: {t}")
            print(f"  Pre-conc value:  ${asset.pre_conc_value:>12,.0f}")
            print(f"  Advance rate:    {asset.advance_rate_label}  ({asset.advance_rate*100:.1f}%)")
            print(f"  Borrowing value: ${asset.borrowing_value:>12,.0f}")
            break

    # Before / After availability
    bw = before.waterfall
    aw = after.waterfall
    print(f"\n{'':>4}{'BEFORE':>16}  {'AFTER':>16}  {'DELTA':>16}")
    print(f"  {'Avail ABV':<20} ${bw.total_abv:>14,.0f}  ${aw.total_abv:>14,.0f}  ${aw.total_abv-bw.total_abv:>+14,.0f}")
    print(f"  {'Excess conc':<20} ${bw.total_excess:>14,.0f}  ${aw.total_excess:>14,.0f}  ${aw.total_excess-bw.total_excess:>+14,.0f}")
    print(f"  {'Net ABV':<20} ${bw.net_abv:>14,.0f}  ${aw.net_abv:>14,.0f}  ${aw.net_abv-bw.net_abv:>+14,.0f}")
    print(f"  {'WAAR':<20} {bw.final_waar*100:>14.2f}%  {aw.final_waar*100:>14.2f}%")
    ba, aa = bw.availability, aw.availability
    print(f"  {'Availability':<20} ${ba.availability:>14,.0f}  ${aa.availability:>14,.0f}  ${aa.availability-ba.availability:>+14,.0f}")
    print()
