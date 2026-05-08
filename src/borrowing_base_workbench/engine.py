"""engine.py — Python calculation backend for the Borrowing Base Workbench.

This module is the execution boundary between the Tkinter frontend (app.py)
and the pure-Python calculation stack (loader → calculator → scenario).

Public API:
  probe_workbook(workbook_path)           — baseline portfolio metrics
  run_pro_forma(workbook_path, scenario)  — before/after pro forma scenario

The workbook is used as a read-only data source only; no Excel runtime,
COM automation, or PowerShell is involved.
"""

from __future__ import annotations

from pathlib import Path

from borrowing_base_workbench.calculator import CONCENTRATION_TESTS, calculate_portfolio
from borrowing_base_workbench.loader import RuntimeData, load_workbook, load_workbook_data
from borrowing_base_workbench.scenario import build_scenario_records, inject_scenario

# _CONCENTRATION_LABEL_ORDER removed — use CONCENTRATION_TESTS from calculator.py
# (imported above). Labels and ordering are now defined in one place.


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
            for lbl, _ in CONCENTRATION_TESTS
        ]

    # Index the waterfall results by label for O(1) lookup
    test_by_label = {t.limit_type: t for t in w.concentration_tests}

    rows = []
    for lbl, _ in CONCENTRATION_TESTS:
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
# Internal helpers
# ---------------------------------------------------------------------------

def _eligibility_for_scenario(calc, scenario_name: str) -> dict:
    """Extract eligibility result for the scenario loan from a PortfolioCalcResult."""
    for asset in calc.assets:
        if asset.obligor_name == scenario_name:
            e = asset.eligibility
            return {"status": "Yes" if e.eligible else "No", "failed_tests": e.failed_tests}
    return {"status": "Yes", "failed_tests": []}


# ---------------------------------------------------------------------------
# Session-aware public API  (preferred — no file I/O after import)
# ---------------------------------------------------------------------------

def import_workbook(workbook_path: str | Path) -> RuntimeData | dict:
    """Load the workbook once into in-memory RuntimeData.

    Returns RuntimeData on success, or an error dict on failure.
    After a successful return the file is no longer required.
    Pass the result to probe_runtime() or run_pro_forma_on_runtime().
    """
    workbook_path = Path(workbook_path)
    if not workbook_path.exists():
        return {
            "status": "error",
            "message": f"Workbook not found: {workbook_path}",
            "remediation": "Check the workbook path in the Admin tab and try again.",
        }
    try:
        return load_workbook(workbook_path)
    except Exception as exc:
        return {
            "status": "error",
            "message": f"Failed to load workbook: {exc}",
            "remediation": "Ensure the workbook is accessible and not exclusively locked.",
        }


def probe_runtime(runtime: RuntimeData) -> dict:
    """Compute baseline portfolio metrics from pre-loaded RuntimeData.

    No file I/O. Uses only in-memory state. Safe to call repeatedly
    without re-reading the workbook.
    """
    data = runtime.to_workbook_data()
    try:
        calc = calculate_portfolio(data)
    except Exception as exc:
        return {"status": "error", "message": f"Calculator error: {exc}"}

    advances = runtime.availability_meta.current_advances
    return {
        "status": "ok",
        "source_path": str(runtime.source_path),
        "imported_at": runtime.imported_at.isoformat(),
        "loader_summary": runtime.summary(),
        "calculator_summary": calc.summary(),
        "metrics": _metrics_from_calculator(calc, advances),
        "concentration_limits": _concentration_limits_from_waterfall(calc),
    }


def run_pro_forma_on_runtime(runtime: RuntimeData, scenario: dict) -> dict:
    """Run a pro forma scenario on pre-loaded RuntimeData.

    No workbook re-read. Uses the cached portfolio state as the baseline,
    injects the scenario loan into an in-memory copy, and returns
    before/after metric snapshots + per-loan eligibility result.

    The original RuntimeData is never mutated.
    """
    data = runtime.to_workbook_data()
    advances = runtime.availability_meta.current_advances

    # --- Before: baseline portfolio -----------------------------------------
    try:
        before_calc = calculate_portfolio(data)
    except Exception as exc:
        return {"status": "error", "message": f"Calculator error (before): {exc}"}

    before_snap = {
        **_metrics_from_calculator(before_calc, advances),
        "concentration_limits": _concentration_limits_from_waterfall(before_calc),
    }

    # --- Build and inject scenario records ----------------------------------
    try:
        scenario_loan, scenario_obligor, scenario_flags = build_scenario_records(scenario)
    except Exception as exc:
        return {"status": "error", "message": f"Scenario construction error: {exc}"}

    scenario_name = scenario_loan.obligor_name
    after_data = inject_scenario(data, scenario_loan, scenario_obligor, scenario_flags)

    # --- After: portfolio with scenario loan --------------------------------
    try:
        after_calc = calculate_portfolio(after_data)
    except Exception as exc:
        return {"status": "error", "message": f"Calculator error (after): {exc}"}

    after_snap = {
        **_metrics_from_calculator(after_calc, advances),
        "concentration_limits": _concentration_limits_from_waterfall(after_calc),
    }

    return {
        "status": "ok",
        "source_path": str(runtime.source_path),
        "before": before_snap,
        "after":  after_snap,
        "eligibility": _eligibility_for_scenario(after_calc, scenario_name),
        "scenario": {
            "sm_support_row": scenario_obligor.source_row,
            "loan_tape_row":  scenario_loan.source_row,
            "portfolio_row":  None,
        },
    }


# ---------------------------------------------------------------------------
# Convenience wrappers (load + run in one call — for backward compat / tests)
# ---------------------------------------------------------------------------

def probe_workbook(workbook_path: str | Path) -> dict:
    """Load workbook and return baseline metrics.

    Convenience wrapper around import_workbook() + probe_runtime().
    Prefer storing the RuntimeData from import_workbook() and calling
    probe_runtime() directly to avoid repeated file reads.
    """
    result = import_workbook(workbook_path)
    if isinstance(result, dict):
        return result
    return probe_runtime(result)


def run_pro_forma(workbook_path: str | Path, scenario: dict) -> dict:
    """Load workbook and run a pro forma scenario.

    Convenience wrapper around import_workbook() + run_pro_forma_on_runtime().
    Prefer pre-loading with import_workbook() when running multiple scenarios.
    """
    result = import_workbook(workbook_path)
    if isinstance(result, dict):
        return result
    return run_pro_forma_on_runtime(result, scenario)
