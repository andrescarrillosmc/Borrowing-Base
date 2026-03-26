"""engine.py — Python backend seam for Excel detachment (v1).

This module is the execution boundary between the Tkinter frontend and the
backend calculation logic.

Phase status:
  - probe_workbook():  loads workbook data via loader.py; metrics are STUBBED
                       (zeroed) until calculator.py is implemented in Phase 3.
  - run_pro_forma():   validates workbook path; metrics are STUBBED until
                       calculator.py is implemented.

The return dict shapes are identical to the retired PowerShell scripts so that
app.py requires no changes beyond the one-line import swap.
"""

from __future__ import annotations

from pathlib import Path

from borrowing_base_workbench.loader import load_workbook_data

# ---------------------------------------------------------------------------
# Concentration limit label order
# ---------------------------------------------------------------------------
# limit_type strings are matched by key in app.py._populate_results_view —
# they must be exactly these strings, in this order.
# ---------------------------------------------------------------------------

_CONCENTRATION_TESTS = [
    {"limit_type": "Max Second Lien & FILO with senior lev >= 1.50x", "limit_percent": 0.20},
    {"limit_type": "Max Second Lien",                                  "limit_percent": 0.10},
    {"limit_type": "Max Non-First Lien",                               "limit_percent": 0.30},
    {"limit_type": "Max EBITDA < $5MM",                                "limit_percent": 0.15},
    {"limit_type": "Max Obligors",                                     "limit_percent": 0.075},
    {"limit_type": "Max Largest Industry",                             "limit_percent": 0.20},
    {"limit_type": "Max Second Largest Industry",                      "limit_percent": 0.15},
    {"limit_type": "Max Other Industries",                             "limit_percent": 0.10},
    {"limit_type": "Fixed Rate",                                       "limit_percent": 0.10},
    {"limit_type": "Max Limited Industry",                             "limit_percent": 0.10},
    {"limit_type": "Max DDTL and Revolver",                            "limit_percent": 0.15},
    {"limit_type": "Max Non-Sponsor/Non-Family Office",                "limit_percent": 0.15},
    {"limit_type": "Max Div Recap Non-Sponsor/Non-Family Office",      "limit_percent": 0.10},
]


def _stub_metrics(current_advances: float = 0.0) -> dict:
    """Zeroed metric placeholders. Replaced by calculator.py in Phase 3."""
    return {
        "availability": 0.0,
        "total_portfolio_par": 0.0,
        "aggregate_adjusted_bv": 0.0,
        "excess_concentration": 0.0,
        "net_adjusted_bv": 0.0,
        "credit_enhancement_test": "",
        "weighted_avg_advance_rate": 0.0,
        "current_advances": current_advances,
    }


def _stub_concentration_limits(policy_limits: list[dict] | None = None) -> list[dict]:
    """13 concentration limit rows with live policy percentages, zeroed actuals."""
    source = policy_limits if policy_limits is not None else _CONCENTRATION_TESTS
    return [
        {
            "limit_type": t["limit_type"],
            "limit_percent": t["limit_percent"],
            "applicable_limit": 0.0,
            "actual": 0.0,
            "excess": 0.0,
        }
        for t in source
    ]


def _concentration_limits_from_policy(policy) -> list[dict]:
    """Build the 13-row concentration limit list from loaded PolicyConfig."""
    pct_map = {
        "Max Second Lien & FILO with senior lev >= 1.50x": policy.conc_max_2l_filo_high_lev,
        "Max Second Lien":                                  policy.conc_max_second_lien,
        "Max Non-First Lien":                               policy.conc_max_non_first_lien,
        "Max EBITDA < $5MM":                                policy.conc_max_small_ebitda,
        "Max Obligors":                                     policy.conc_max_obligor,
        "Max Largest Industry":                             policy.conc_max_largest_industry,
        "Max Second Largest Industry":                      policy.conc_max_second_industry,
        "Max Other Industries":                             policy.conc_max_other_industries,
        "Fixed Rate":                                       policy.conc_fixed_rate,
        "Max Limited Industry":                             policy.conc_limited_industry,
        "Max DDTL and Revolver":                            policy.conc_max_ddtl_revolver,
        "Max Non-Sponsor/Non-Family Office":                policy.conc_max_non_sponsor,
        "Max Div Recap Non-Sponsor/Non-Family Office":      policy.conc_max_div_recap,
    }
    return [
        {
            "limit_type": t["limit_type"],
            "limit_percent": pct_map.get(t["limit_type"], t["limit_percent"]),
            "applicable_limit": 0.0,
            "actual": 0.0,
            "excess": 0.0,
        }
        for t in _CONCENTRATION_TESTS
    ]


# ---------------------------------------------------------------------------
# Public API
# ---------------------------------------------------------------------------

def probe_workbook(workbook_path: str | Path) -> dict:
    """Read the current baseline state of the portfolio from the workbook.

    Replaces: excel_runner.probe_excel_workbook(workbook_path, script_path)

    Phase 2 status: loads workbook data via loader.py; populates policy-driven
    concentration limit percentages and current_advances from the workbook.
    All computed metrics (availability, ABV, etc.) remain STUBBED at 0.0
    until calculator.py is implemented in Phase 3.

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

    metrics = _stub_metrics(current_advances=data.availability_meta.current_advances)
    concentration_limits = _concentration_limits_from_policy(data.policy)

    return {
        "status": "ok",
        "workbook_path": str(workbook_path),
        "loader_summary": data.summary(),  # extra field — ignored by frontend, useful for debug
        "metrics": metrics,
        "concentration_limits": concentration_limits,
    }


def run_pro_forma(workbook_path: str | Path, scenario: dict) -> dict:
    """Run a pro forma scenario and return before/after results.

    Replaces: excel_runner.run_pro_forma_workbook(workbook_path, script_path, scenario)

    Phase 2 status: loads workbook data; all computed metrics STUBBED at 0.0
    until calculator.py is implemented in Phase 3.

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

    concentration_limits = _concentration_limits_from_policy(data.policy)
    advances = data.availability_meta.current_advances

    before = {**_stub_metrics(advances), "concentration_limits": concentration_limits}
    after  = {**_stub_metrics(advances), "concentration_limits": concentration_limits}

    return {
        "status": "ok",
        "workbook_path": str(workbook_path),
        "before": before,
        "after": after,
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
