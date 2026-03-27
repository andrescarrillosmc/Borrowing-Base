# Architecture — Borrowing Base Workbench

## Overview

The application is a Python/Tkinter desktop tool that reads Star Mountain Capital's
BDC borrowing base Excel workbook as a **read-only data source** and performs all
calculations in pure Python.

No Excel runtime, COM automation, or PowerShell is required at runtime.

---

## Module Map

```
src/borrowing_base_workbench/
  app.py         Tkinter UI — admin panel, scenario form, results display
  engine.py      Backend boundary — public API called by app.py
  loader.py      Workbook reader — openpyxl, no Excel runtime
  calculator.py  Portfolio calculation engine
  scenario.py    Pro forma record construction
  models.py      (re-exported via loader) shared data classes
  validation.py  Pre-flight checks and commentary
  analysis.py    Workbook structure diagnosis
```

---

## Workbook Reader (`loader.py`)

**Input:** path to `.xlsm` / `.xlsx` workbook
**Output:** `WorkbookData` dataclass

Reads six workbook sheets using `openpyxl` (data_only=True):

| Sheet | Loaded as |
|-------|-----------|
| `Loan Tape - Settled` | `list[LoanRecord]` |
| `SM Support` | `list[ObligorRecord]` |
| `VAE Log` | `list[VaeRecord]` |
| `AGENT` | `PolicyConfig` (advance rates, concentration limits, WAAR caps, collateral tiers) |
| `CL` | Manual portfolio flags (`dict[str, ManualPortfolioFlags]`) |
| `Availability` | `AvailabilityMeta` (facility amount, current advances) |

The workbook is opened once per probe/scenario call; no state is retained.

---

## Calculation Engine (`calculator.py`)

**Input:** `WorkbookData`
**Output:** `PortfolioCalcResult`

### Per-asset calculation (`calculate_asset`)

For each `LoanRecord`:

1. **Eligibility** (`EligibilityResult.compute`) — 10 binary tests:
   - A: denomination is USD
   - B: country is US
   - C: loan type is in approved set
   - D: not cov-lite
   - E: PIK < 50%
   - F: purchase price >= minimum
   - G: EBITDA >= minimum (Second Lien / FILO only)
   - H: initial total leverage within policy limit
   - I: maturity <= 7 years from inclusion date
   - J: not flagged as DIP or bankrupt VAE

2. **Collateral tier** (`_lookup_collateral_tier`) — maps leverage to a tier percentage:
   - First Lien / DDTL / Revolver: uses `net_detachment`
   - FILO / Second Lien: uses `net_attachment`

3. **Advance rate** (`_pick_advance_rate`) — selects the applicable rate from `PolicyConfig`
   based on loan type and EBITDA bucket

4. **Pre-concentration value** — `min(fmv, tier_pct × olb) × advance_rate`
   - VAE-affected assets: `vae_assigned_value` replaces `fmv` in the min
   - Ineligible assets: `pre_conc_value = 0`

### Portfolio waterfall (`calculate_portfolio`)

1. Sum all `pre_conc_value` → `total_pre_conc_eligible_value` (Aggregate ABV)

2. **WAAR** — weighted average advance rate across eligible assets, capped by
   obligor-count diversity:
   - ≥ 20 obligors: cap = 65%
   - 12–19 obligors: cap = 62.5%
   - < 12 obligors: cap = 55%

3. **Concentration tests** — 13 tests, each computes `qualifying_value` against
   `applicable_limit = limit_pct × net_abv`; excess = max(0, qualifying − limit)

4. **Net ABV** = total_pre_conc_eligible_value − sum(all excess concentration)

5. **Availability** = min(Test A, Test B, Test C):
   - Test A: facility cap ($200M)
   - Test B: Net ABV × final_WAAR
   - Test C: Net ABV × 0.90 − current_advances (credit enhancement)

---

## Scenario Execution (`scenario.py` + `engine.py`)

**Input:** 23-key string dict from the scenario form
**Output:** `{before, after, eligibility}` — before/after metric snapshots

### Flow

```
build_scenario_records(scenario_dict)
  → (LoanRecord, ObligorRecord, ManualPortfolioFlags)

inject_scenario(baseline_data, loan, obligor, flags)
  → new WorkbookData (immutable copy — original untouched)

calculate_portfolio(augmented_data)
  → after_calc
```

### Leverage derivation

The scenario form collects financial inputs; leverage is derived to mirror
the SM Support sheet formula-proxy columns:

```
net_detachment = (drawn_revolver + first_out - cash_balance) / ltm_adj_ebitda
net_attachment = (drawn_revolver + first_out + pari_passu + total_sm - cash_balance) / ltm_adj_ebitda
```

`net_detachment` drives collateral tier for First Lien loans.
`net_attachment` drives eligibility Test H and FILO/2L tier selection.

---

## Frontend / Backend Boundary (`engine.py`)

`app.py` calls exactly two functions:

```python
from borrowing_base_workbench.engine import (
    probe_workbook as probe_excel_workbook,
    run_pro_forma as run_pro_forma_workbook,
)
```

### `probe_workbook(workbook_path) -> dict`

Returns:
```python
{
    "status": "ok",
    "metrics": {
        "availability": float,
        "aggregate_adjusted_bv": float,
        "net_adjusted_bv": float,
        "weighted_avg_advance_rate": float,
        "excess_concentration": float,
        # ... 10 more fields
    },
    "concentration_limits": [ {limit_type, limit_percent, applicable_limit, actual, excess}, ... ]
}
```

### `run_pro_forma(workbook_path, scenario) -> dict`

Returns:
```python
{
    "status": "ok",
    "before": { <same shape as probe_workbook metrics> },
    "after":  { <same shape as probe_workbook metrics> },
    "eligibility": {
        "status": "Yes" | "No",
        "failed_tests": [list of test names]
    }
}
```

---

## Data Flow Diagram

```
Excel Workbook (.xlsm)
    |
    | openpyxl (data_only=True)
    v
loader.py  ->  WorkbookData
                    |
                    +----------------------------+
                    |                            |
                    v                            v
              calculate_portfolio()       inject_scenario()
              (baseline)                  -> augmented WorkbookData
                    |                            |
                    v                            v
              before_metrics           calculate_portfolio()
                                             (after)
                                                 |
                                                 v
                                          after_metrics
                                        + eligibility result
                                                 |
                                                 v
                                           engine.py API
                                                 |
                                                 v
                                            app.py UI
```

---

## Key Constraints

- The workbook is **never written to** by the Python code
- Scenario injection is **in-memory only** — no file is created
- All calculations run in a single thread (Tkinter + synchronous calls)
- Integration tests require the workbook at the path in `tests/test_engine.py:WORKBOOK`
