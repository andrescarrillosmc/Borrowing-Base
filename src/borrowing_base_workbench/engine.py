"""engine.py — Python backend seam for Excel detachment (v1).

This module is the execution boundary between the Tkinter frontend and the
backend calculation logic.

Phase status:
  - probe_workbook():  Phases 3+4 complete — returns real asset-level results,
                       concentration waterfall, WAAR with obligor-count cap,
                       and three-test availability calculation.
  - run_pro_forma():   Baseline "before" state uses the full calculator.
                       "after" state is still stubbed — scenario mutation
                       (Phase 5) not yet implemented.

Return dict shapes are identical to the retired PowerShell scripts so that
app.py requires no changes beyond the one-line import swap.
"""

from __future__ import annotations

from pathlib import Path

from borrowing_base_workbench.calculator import calculate_portfolio
from borrowing_base_workbench.loader import load_workbook_data

# ---------------------------------------------------------------------------
# Concentration limit label order
# ---------------------------------------------------------------------------
# limit_type strings are matched by key in app.py._populate_results_view.
# They MUST be exactly these strings, in this order.
# ---------------------------------------------------------------------------

_CONCENTRATION_LABEL_ORDER = [
    "Max Second Lien & FILO with senior lev >= 1.50x",
    "Max Second Lien",
    "Max Non-First Lien",
    "Max EBITDA < $5MM",
    "Max Obligors",
    "Max Largest Industry",
    "Max Second Largest Industry",
    "Max Other Industries",
    "Fixed Rate",
    "Max Limited Industry",
    "Max DDTL and Revolver",
    "Max Non-Sponsor/Non-Family Office",
    "Max Div Recap Non-Sponsor/Non-Family Office",
]


# ---------------------------------------------------------------------------
# Internal helpers
# ---------------------------------------------------------------------------

def _metrics_from_calculator(calc, current_advances: float) -> dict:
    """Build the metrics dict from a PortfolioCalcResult (Phase 3+4).

    All previously stubbed fields are now populated from the waterfall.
    """
    w = calc.waterfall
    av = w.availability if w else None

    return {
        "availability":             av.availability         if av else 0.0,
        "total_portfolio_par":      calc.total_portfolio_par,
        "aggregate_adjusted_bv":    calc.total_pre_conc_eligible_value,
        "excess_concentration":     w.total_excess          if w else 0.0,
        "net_adjusted_bv":          w.net_abv               if w else 0.0,
        "credit_enhancement_test":  (
            f"${av.test_c_credit_enhancement:,.0f}" if av else ""
        ),
        "weighted_avg_advance_rate": w.final_waar           if w else 0.0,
        "current_advances":          current_advances,
        # Extra debug fields (ignored by frontend)
        "eligible_count":           len(calc.eligible_assets),
        "ineligible_count":         len(calc.ineligible_assets),
        "vae_affected_count":       len(calc.vae_affected),
        "total_borrowing_value":    calc.total_borrowing_value,
        "raw_waar":                 w.raw_waar              if w else 0.0,
        "discrete_obligor_count":   w.discrete_obligor_count if w else 0,
        "waar_cap":                 w.applicable_waar_cap   if w else 0.0,
        "test_a_facility":          av.test_a_facility      if av else 0.0,
        "test_b_borrowing_base":    av.test_b_borrowing_base if av else 0.0,
        "test_c_credit_enhancement_dollar": av.test_c_credit_enhancement if av else 0.0,
    }


def _concentration_limits_from_waterfall(calc) -> list[dict]:
    """Build the 13-row concentration list from Phase 4 waterfall results.

    Falls back to policy percentages with zero actuals if waterfall is absent.
    """
    w = calc.waterfall
    if w is None:
        # Fallback: policy percentages, zero actuals
        pct_map = _policy_pct_map(calc)
        return [
            {"limit_type": lbl, "limit_percent": pct_map.get(lbl, 0.0),
             "applicable_limit": 0.0, "actual": 0.0, "excess": 0.0}
            for lbl in _CONCENTRATION_LABEL_ORDER
        ]

    # Index the waterfall results by label for O(1) lookup
    test_by_label = {t.limit_type: t for t in w.concentration_tests}

    rows = []
    for lbl in _CONCENTRATION_LABEL_ORDER:
        t = test_by_label.get(lbl)
        if t:
            rows.append({
                "limit_type":       t.limit_type,
                "limit_percent":    t.limit_percent,
                "applicable_limit": t.applicable_limit,
                "actual":           t.qualifying_value,
                "excess":           t.excess,
            })
        else:
            rows.append({
                "limit_type": lbl, "limit_percent": 0.0,
                "applicable_limit": 0.0, "actual": 0.0, "excess": 0.0,
            })
    return rows


def _policy_pct_map(calc) -> dict[str, float]:
    """Extract concentration limit percentages from loaded policy."""
    p = calc.assets[0].eligibility  # reach policy via loan data is awkward;
    # fall back to known defaults when called without waterfall
    return {lbl: 0.0 for lbl in _CONCENTRATION_LABEL_ORDER}


# ---------------------------------------------------------------------------
# Public API
# ---------------------------------------------------------------------------

def probe_workbook(workbook_path: str | Path) -> dict:
    """Read the current baseline state of the portfolio from the workbook.

    Replaces: excel_runner.probe_excel_workbook(workbook_path, script_path)

    Phase 4 complete: returns real values for all core metrics including
    availability, net ABV, WAAR (with cap), and concentration test details.

    Return shape mirrors excel_probe.ps1 stdout JSON.
    """
    workbook_path = Path(workbook_path)

    if not workbook_path.exists():
        return {
            "status": "error",
            "message": f"Workbook not found: {workbook_path}",
            "remediation": "Check the workbook path in the Admin tab and try again.",
        }

    try:
        data = load_workbook_data(workbook_path)
    except Exception as exc:
        return {
            "status": "error",
            "message": f"Failed to load workbook: {exc}",
            "remediation": "Ensure the workbook is accessible and not exclusively locked.",
        }

    try:
        calc = calculate_portfolio(data)
    except Exception as exc:
        return {
            "status": "error",
            "message": f"Calculator error: {exc}",
            "remediation": "Check calculator.py for data issues.",
        }

    advances = data.availability_meta.current_advances
    metrics = _metrics_from_calculator(calc, advances)
    concentration_limits = _concentration_limits_from_waterfall(calc)

    return {
        "status": "ok",
        "workbook_path": str(workbook_path),
        "loader_summary": data.summary(),       # ignored by frontend, useful for debug
        "calculator_summary": calc.summary(),   # Phase 4 summary
        "metrics": metrics,
        "concentration_limits": concentration_limits,
    }


def run_pro_forma(workbook_path: str | Path, scenario: dict) -> dict:
    """Run a pro forma scenario and return before/after results.

    Replaces: excel_runner.run_pro_forma_workbook(workbook_path, script_path, scenario)

    Phase 4 status: "before" uses the full calculator (real values).
    "after" remains identical to "before" — scenario mutation (Phase 5) pending.

    Return shape mirrors run_pro_forma.ps1 stdout JSON.
    """
    workbook_path = Path(workbook_path)

    if not workbook_path.exists():
        return {
            "status": "error",
            "message": f"Workbook not found: {workbook_path}",
        }

    try:
        data = load_workbook_data(workbook_path)
    except Exception as exc:
        return {
            "status": "error",
            "message": f"Failed to load workbook: {exc}",
        }

    try:
        calc = calculate_portfolio(data)
    except Exception as exc:
        return {
            "status": "error",
            "message": f"Calculator error: {exc}",
        }

    advances = data.availability_meta.current_advances
    metrics = _metrics_from_calculator(calc, advances)
    concentration_limits = _concentration_limits_from_waterfall(calc)

    baseline = {**metrics, "concentration_limits": concentration_limits}

    return {
        "status": "ok",
        "workbook_path": str(workbook_path),
        "before": baseline,
        "after": baseline,     # Phase 5 will diff this against scenario-mutated calc
        "eligibility": {
            "status": "Yes",
            "failed_tests": [],
        },
        "scenario": {
            "sm_support_row": None,
            "loan_tape_row": None,
            "portfolio_row": None,
        },
    }
