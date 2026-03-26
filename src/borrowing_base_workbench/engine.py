"""engine.py — Python backend seam for Excel detachment (v1).

This module is the execution boundary between the Tkinter frontend and the
backend calculation logic.  In this initial phase the two public functions
return deterministic stub responses that satisfy every UI rendering path in
app.py.  No Excel, no PowerShell, no subprocess.

Replacement target:
    app.py previously imported from excel_runner.py:
        probe_excel_workbook(workbook_path, script_path) -> dict
        run_pro_forma_workbook(workbook_path, script_path, scenario) -> dict

    app.py now imports from this module:
        probe_workbook(workbook_path) -> dict
        run_pro_forma(workbook_path, scenario) -> dict

    The script_path argument is gone; it was only needed to locate the
    PowerShell file.  The return dict shapes are identical to the PS scripts.

Next phase: replace the stub bodies with loader.py + calculator.py calls.
"""

from __future__ import annotations

from pathlib import Path

# ---------------------------------------------------------------------------
# Concentration limit metadata
# ---------------------------------------------------------------------------
# The 13 tests in the exact order expected by the frontend (limit_type strings
# are matched by key in app.py._populate_results_view).
# limit_percent values come from AGENT!D57-D69.
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


def _stub_metrics() -> dict:
    """Return zeroed-out metric placeholders matching the probe output shape."""
    return {
        "availability": 0.0,
        "total_portfolio_par": 0.0,
        "aggregate_adjusted_bv": 0.0,
        "excess_concentration": 0.0,
        "net_adjusted_bv": 0.0,
        "credit_enhancement_test": "",
        "weighted_avg_advance_rate": 0.0,
        "current_advances": 0.0,
    }


def _stub_concentration_limits() -> list[dict]:
    """Return 13 zeroed concentration limit rows matching the probe output shape."""
    return [
        {
            "limit_type": t["limit_type"],
            "limit_percent": t["limit_percent"],
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

    Current implementation: stub.  Returns a valid response shape with zeroed
    numeric values so that the Results and Admin tabs render without errors.

    Next phase: load WorkbookData via loader.py, run calculator.py on the
    baseline portfolio, and return the computed values.

    Return shape mirrors excel_probe.ps1 stdout JSON:
        {
            "status": "ok",
            "metrics": { 8 keys },
            "concentration_limits": [ 13 dicts ]
        }
    """
    workbook_path = Path(workbook_path)

    if not workbook_path.exists():
        return {
            "status": "error",
            "message": f"Workbook not found: {workbook_path}",
            "remediation": "Check the workbook path in the Admin tab and try again.",
        }

    return {
        "status": "ok",
        "workbook_path": str(workbook_path),
        "metrics": _stub_metrics(),
        "concentration_limits": _stub_concentration_limits(),
    }


def run_pro_forma(workbook_path: str | Path, scenario: dict) -> dict:
    """Run a pro forma scenario against the workbook and return before/after results.

    Replaces: excel_runner.run_pro_forma_workbook(workbook_path, script_path, scenario)

    Current implementation: stub.  Returns a valid response shape with zeroed
    numeric values and a passing eligibility result so that all Results tab
    sections render without errors.

    Next phase: load WorkbookData via loader.py, build a synthetic LoanRecord
    from `scenario`, run calculator.py on baseline (before) and on
    baseline + scenario loan (after), and return the delta.

    Return shape mirrors run_pro_forma.ps1 stdout JSON:
        {
            "status": "ok",
            "before": { metrics + concentration_limits },
            "after":  { metrics + concentration_limits },
            "eligibility": { "status": str, "failed_tests": list },
            "scenario": { row numbers (informational) }
        }
    """
    workbook_path = Path(workbook_path)

    if not workbook_path.exists():
        return {
            "status": "error",
            "message": f"Workbook not found: {workbook_path}",
        }

    before = {**_stub_metrics(), "concentration_limits": _stub_concentration_limits()}
    after  = {**_stub_metrics(), "concentration_limits": _stub_concentration_limits()}

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
