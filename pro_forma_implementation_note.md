# Pro Forma Scenario Execution — Implementation Note

**Branch:** `engine-detach-v1`
**Phase:** 5 of 6 (Scenario Mutation)

---

## What Was Built

### New file: `src/borrowing_base_workbench/scenario.py`

Two public functions:

- **`build_scenario_records(scenario: dict[str, str])`** — converts the 23-key
  form dict into a `(LoanRecord, ObligorRecord, ManualPortfolioFlags)` triple
  ready for injection.

- **`inject_scenario(data, loan, obligor, flags)`** — returns a new
  `WorkbookData` with the synthetic records appended.  Uses `dataclasses.replace()`
  — no mutation of the original workbook data.

### Updated: `src/borrowing_base_workbench/engine.py`

`run_pro_forma()` is now fully implemented:

1. Loads baseline `WorkbookData` from the workbook.
2. Runs `calculate_portfolio()` → **before** metrics (real values, same as probe).
3. Calls `build_scenario_records(scenario)` → synthetic records.
4. Calls `inject_scenario(data, ...)` → augmented in-memory copy.
5. Runs `calculate_portfolio()` on augmented data → **after** metrics.
6. Extracts per-loan eligibility for the scenario company from the after calc.
7. Returns `{status, before, after, eligibility}` — exact frontend contract.

---

## Scenario Field Mapping

| Form key | Type | → Record field |
|----------|------|----------------|
| `company_name` | str | `LoanRecord.obligor_name`, `ObligorRecord.obligor_name` |
| `security_type` | str ("First Lien" / "FILO") | `LoanRecord.loan_type` |
| `ltm_revenue` | currency | `ObligorRecord.ltm_revenue` (dollars) |
| `ltm_adj_ebitda` | currency | `ObligorRecord.ltm_adj_ebitda` (dollars); `LoanRecord.current_ebitda_mm` = /1M |
| `drawn_revolver` | currency | `ObligorRecord.drawn_revolver` |
| `first_out_balance` | currency | `ObligorRecord.first_out` |
| `pari_passu` | currency | `ObligorRecord.pari_passu` |
| `bdc_balance` | currency | `ObligorRecord.debt_balance`, `LoanRecord.olb`, `LoanRecord.commitment_balance` |
| `total_sm_balance` | currency | `ObligorRecord.total_sm_balance` |
| `cash_balance` | currency | `ObligorRecord.cash_balance` |
| `interest_coverage` | multiple | `ObligorRecord.interest_coverage` |
| `loan_denomination` | str | `LoanRecord.denomination` (default: "USD") |
| `purchase_price` | fraction | `LoanRecord.purchase_price`; FMV = purchase_price × bdc_balance |
| `country` | str | `LoanRecord.country` (default: "United States") |
| `investment_date` | date | `LoanRecord.investment_date`, `.inclusion_date`, `.inception_date` |
| `maturity_date` | date | `LoanRecord.maturity_date` |
| `rate_type` | str | `LoanRecord.rate_type` ("Floating" / "Fixed") |
| `industry_classification` | str | `LoanRecord.industry` |
| `payment_frequency` | str | `LoanRecord.payment_frequency` ("M" / "Q") |
| `pik_pct` | fraction | `LoanRecord.pik_pct` |
| `spread` | fraction | `LoanRecord.spread` |
| `sofr_floor` | fraction | `LoanRecord.sofr_floor` |
| `attach_point` | multiple | Fallback for net_detachment when EBITDA = 0 |

---

## Leverage Derivation

Mirrors SM Support formula-proxy columns M and N:

```
net_detachment = (drawn_revolver + first_out - cash_balance) / ltm_adj_ebitda
net_attachment = (drawn_revolver + first_out + pari_passu + total_sm - cash) / ltm_adj_ebitda
```

`net_detachment` drives FILO advance-rate bucketing.
`net_attachment` is used as `initial_total_leverage` for eligibility Test H.

If EBITDA = 0 (degenerate input), `attach_point` is used as a direct proxy.

---

## Defaults for Synthetic Loans

| Field | Value |
|-------|-------|
| `cov_lite` | `False` |
| `pik_pct` | from form (typically 0) |
| `is_ddtl` | `True` only if loan_type == "DDTL" |
| `unfunded` | `0.0` |
| `limited_industry` | `False` |
| `non_sponsor` | `False` |
| `div_recap` | `False` |
| `is_dip` | `False` |
| `inclusion_date` | same as `investment_date` (no prior workbook history) |
| Mark (low/high/avg) | `purchase_price` (at-par for new originations) |

---

## Observed Results (2025-02-11 workbook — example scenario)

**Scenario:** Acme Corp, First Lien, $25MM EBITDA, $20MM OLB, at par

| Metric | Before | After | Delta |
|--------|--------|-------|-------|
| Aggregate ABV | $120,210,997 | $140,210,997 | +$20,000,000 |
| Excess concentration | $29,281,364 | $27,829,792 | −$1,451,572 |
| Net ABV | $90,929,633 | $112,381,205 | +$21,451,572 |
| WAAR | 60.80% | 61.75% | — |
| Availability | $60,753,941 | $74,870,393 | +$14,116,452 |

The excess concentration decreases because adding a new First Lien obligor
distributes the portfolio more broadly, relaxing the per-obligor cap test.

---

## Testing

### Direct CLI (scenario.py)
```bash
cd C:\Users\henry.yan\borrowing-base
PYTHONPATH=src python src\borrowing_base_workbench\scenario.py
```
Prints: scenario loan eligibility, before/after table with delta column.

### Via engine
```python
from borrowing_base_workbench.engine import run_pro_forma

scenario = {
    "company_name": "Acme Corp",
    "security_type": "First Lien",
    "ltm_adj_ebitda": "25000000",
    "bdc_balance": "20000000",
    "total_sm_balance": "20000000",
    "purchase_price": "1.0",
    "investment_date": "2025-01-15",
    "maturity_date": "2031-01-15",
    "industry_classification": "Software",
    "loan_denomination": "USD",
    "country": "United States",
    "rate_type": "Floating",
    # ... remaining keys can be empty strings for optional fields
}

r = run_pro_forma(
    r"C:\Users\henry.yan\Downloads\2025-02-11_BDC Borrowing_Base_v8.xlsm",
    scenario,
)
print(r["eligibility"])                                 # {"status": "Yes", "failed_tests": []}
print(r["after"]["availability"])                       # 74,870,393
print(r["after"]["availability"] - r["before"]["availability"])  # +14,116,452
```

---

## What Remains for Phase 6

1. **Unfunded equity deduction** — implement L37 array formula if needed
   (currently ~$0 for this portfolio; impact negligible)
2. **Write-back stubs** — `scenario.sm_support_row` / `loan_tape_row` return
   the sentinel `9999`; a future Phase could write the scenario loan back to
   a staged workbook copy if Excel-side output is needed
3. **Rate type "Fixed" WAAR impact** — Fixed Rate loans change the composition
   of the Fixed Rate concentration test; currently tracked correctly in the
   waterfall via `LoanRecord.rate_type`
