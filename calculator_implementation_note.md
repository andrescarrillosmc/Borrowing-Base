# Calculator — Implementation Note

**Branch:** `engine-detach-v1`
**Phase:** 3 of 6 (Asset-Level Calculation Engine)

---

## What Was Built

### New file: `src/borrowing_base_workbench/calculator.py`

Asset-level borrowing base logic.  Takes `WorkbookData` (from `loader.py`) and
returns typed calculation results.  No Excel COM, no PowerShell, no cached formula reads.

Public entry point:
```python
from borrowing_base_workbench.calculator import calculate_portfolio

result = calculate_portfolio(data)   # data = WorkbookData from loader
print(result.summary())
```

### Updated: `src/borrowing_base_workbench/engine.py`

`probe_workbook()` now calls `calculate_portfolio()` and returns real values for:
- `total_portfolio_par` — sum of OLB for all 41 loans
- `aggregate_adjusted_bv` — sum of pre-concentration assigned values (eligible loans)
- `weighted_avg_advance_rate` — implied WAAR (pre-concentration, pre-obligor-count cap)
- `eligible_count`, `ineligible_count`, `vae_affected_count` — debug fields

Still stubbed (Phase 4): `availability`, `excess_concentration`, `net_adjusted_bv`,
`credit_enhancement_test`.

---

## Phase 3 Calculation Logic

### 1. Eligibility (10 tests per loan → `eligible: bool`)

| Test | Field(s) Used | Fail Condition |
|------|--------------|----------------|
| A: Valid type | `loan.loan_type` | Not in {First Lien, FILO, Second Lien, DDTL, Revolver} |
| B: USD denomination | `loan.denomination` | Not "USD" |
| C: US domicile | `loan.country` | Not in US country set |
| D: Not cov-lite | `loan.cov_lite` | cov_lite == True |
| E: Non-PIK/DIP | `loan.pik_pct`, `flags.is_dip` | pik_pct > 0 OR is_dip == True |
| F: Acquisition price | `loan.purchase_price` | < `policy.min_acquisition_price` |
| G: 2L EBITDA (if 2L) | `loan.current_ebitda_mm` | < `policy.min_ebitda_second_lien_mm` (N/A otherwise) |
| H: Total leverage | `loan.initial_total_leverage` | > `policy.max_total_leverage_at_inclusion` (pass if None) |
| I: Original maturity | `loan.maturity_date`, `loan.inclusion_date` | (maturity − inclusion) / 365 > `policy.max_original_maturity_years` |
| J: Not defaulted | VAE `event_type` keywords | Has VAE with "bankrupt" / "chapter 11" / "chapter 7" in event type |

**Note on Test I:** Uses `inclusion_date` (col AO), **not** `investment_date` (col U).
These differ for many loans in the current workbook.

**Note on Test E:** `is_dip` comes from `ManualPortfolioFlags.is_dip` (Portfolio col CA),
not from the Loan Tape.

### 2. Collateral Tier Lookup

Four tiers from `PolicyConfig.collateral_tiers` (AGENT rows 49–52):

| Tier | % | Leverage threshold |
|------|---|--------------------|
| 1 | 100% | FL net det ≤ 3.25x / FILO-2L net att ≤ 3.75x |
| 2 | 92.5% | FL ≤ 4.00x / FILO-2L ≤ 4.25x |
| 3 | 85% | FL ≤ 5.00x / FILO-2L ≤ 5.00x |
| 4 | 0% | > 5.00x (sentinel appended by loader) |

Leverage metric used:
- First Lien / DDTL / Revolver → `loan.net_detachment` (= SM Support col M)
- FILO / Second Lien → `loan.net_attachment` (= SM Support col N)

`collateral_tier_value = tier_pct × loan.olb`

### 3. VAE Override → Pre-Concentration Assigned Value

For eligible loans:
```
candidates = [loan.fmv, collateral_tier_value]
if most_recent_vae.vae_agent_assigned_value is not None:
    candidates.append(vae_agent_assigned_value)
pre_conc_value = min(candidates)
```

Most recent VAE = max by `vae_date` among VAEs with non-None `vae_agent_assigned_value`.

For ineligible loans: `pre_conc_value = 0.0`

Many VAE loans in the current workbook have `vae_agent_assigned_value = $1` (agent's
nominal write-down) or `$0` (complete write-off).  This is correct behavior — those
loans contribute near-zero to the borrowing base.

### 4. Advance Rate Bucketing

| Loan type | Metric | Bucket | Rate |
|-----------|--------|--------|------|
| First Lien / DDTL / Revolver | `current_ebitda_mm` | > $20MM | `policy.rate_first_lien_large` |
| | | $10–20MM | `policy.rate_first_lien_mid` |
| | | < $10MM | `policy.rate_first_lien_small` |
| FILO | `net_detachment` | < 0.75x | `policy.rate_filo_low_lev` |
| | | 0.75–1.25x | `policy.rate_filo_mid_lev` |
| | | > 1.25x | `policy.rate_filo_high_lev` |
| Second Lien | — | fixed | `policy.rate_second_lien` |

`borrowing_value = advance_rate × pre_conc_value`

---

## Result Dataclasses

### `EligibilityResult`
Ten boolean test fields + `eligible: bool` + `failed_tests: list[str]` (human-readable reasons).

### `AssetCalcResult` — per loan
All raw source values used, plus: eligibility, collateral tier (index + pct + value),
VAE info, `pre_conc_value`, advance rate (label + float), `borrowing_value`.

### `PortfolioCalcResult` — aggregate
List of `AssetCalcResult` with computed properties:
- `eligible_assets`, `ineligible_assets`, `vae_affected`
- `total_portfolio_par` — sum of all OLBs
- `total_pre_conc_eligible_value` — sum of pre_conc_value for eligible assets
- `total_borrowing_value` — sum of advance_rate × pre_conc_value
- `implied_weighted_avg_advance_rate` — pre-concentration, pre-obligor-count cap

---

## Observed Results (2025-02-11 workbook, measurement date 2025-12-31)

| Metric | Value |
|--------|-------|
| Total assets | 41 |
| Eligible | 25 |
| Ineligible | 16 (mostly PIK loans) |
| VAE-affected eligible | 11 |
| Total portfolio par (OLB) | $330,231,191 |
| Pre-conc eligible value | $120,210,997 |
| Pre-conc borrowing value | $73,085,307 |
| Implied WAAR (pre-cap) | 60.80% |
| Current advances | $122,300,000 |

**Ineligibility breakdown:** 14 of 16 fail Test E (PIK >0%). 2 fail Test I (maturity
just over 7 years). No loans fail leverage, domicile, or denomination tests.

**Collateral tier distribution (eligible):** 21 in Tier 1 (100%), 1 in Tier 3 (85%),
3 in Tier 4 (0% — leverage > 5x, contributing $0 pre-conc value despite being
technically eligible on all other tests).

---

## Testing

### Direct CLI test
```bash
cd C:\Users\henry.yan\borrowing-base
PYTHONPATH=src python src\borrowing_base_workbench\calculator.py
```
Prints: eligible/ineligible counts, pre-conc values, VAE detail, advance rate distribution,
tier distribution.

### Via engine
```python
from borrowing_base_workbench.engine import probe_workbook
r = probe_workbook(r"C:\Users\henry.yan\Downloads\2025-02-11_BDC Borrowing_Base_v8.xlsm")
print(r["status"])                         # "ok"
print(r["metrics"]["eligible_count"])      # 25
print(r["metrics"]["aggregate_adjusted_bv"])   # 120,210,997
print(r["metrics"]["weighted_avg_advance_rate"])  # 0.6080
print(r["calculator_summary"])
```

---

## What Remains for Phase 4 (concentration waterfall)

1. **Concentration tests** — 13 sequential haircut tests → per-test excess amounts
2. **Weighted-average advance rate cap** — based on obligor count vs. thresholds
3. **Net adjusted borrowing value** — aggregate_adjusted_bv − excess_concentration
4. **Availability** — MIN(facility_amount, net_abv, credit_enhancement_test)
5. **Scenario loan construction** — build `LoanRecord + ObligorRecord` from scenario dict,
   rerun `calculate_portfolio()` on augmented data, return before/after deltas
