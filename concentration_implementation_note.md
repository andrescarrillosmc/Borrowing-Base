# Concentration Waterfall & Availability — Implementation Note

**Branch:** `engine-detach-v1`
**Phase:** 4 of 6 (Concentration Waterfall + Availability)

---

## What Was Built

### Extended: `src/borrowing_base_workbench/calculator.py`

Three new dataclasses added:
- **`ConcentrationTestResult`** — per-test: qualifying_value, applicable_limit, excess
- **`AvailabilityResult`** — three-test availability (L22 / L23 / L25 / L26 mirror)
- **`WaterfallResult`** — full Phase 4 output: concentration tests + WAAR cap + availability

Five concentration-classification fields added to **`AssetCalcResult`**:
`industry`, `is_fixed_rate`, `limited_industry`, `non_sponsor`, `div_recap`

New function `_run_waterfall()`: implements the 13-test sequential concentration
waterfall, WAAR calculation with obligor-count cap, and the three-test availability.

`calculate_portfolio()` now calls `_run_waterfall()` and stores result on
`PortfolioCalcResult.waterfall`.

### Updated: `src/borrowing_base_workbench/engine.py`

`probe_workbook()` now returns real values for all previously stubbed core outputs:
- `availability` — Availability!L26
- `excess_concentration` — CL!M44
- `net_adjusted_bv` — CL!M45
- `weighted_avg_advance_rate` — Availability!L48 (WAAR with obligor-count cap)
- `credit_enhancement_test` — formatted Test C dollar value

`_concentration_limits_from_waterfall()` builds the 13-row concentration limits list
from live waterfall results (qualifying_value, applicable_limit, excess per test).

`run_pro_forma()` now uses the full calculator for the "before" state (real values).
"after" state remains identical — scenario mutation is Phase 5.

---

## Concentration Waterfall Logic

### Sequential vs parallel

The waterfall is **sequential**: each test operates on `cur_val[loan_row]`, which
is the value remaining after all prior tests' haircuts. When a test generates
excess, it is distributed proportionally across qualifying loans, reducing their
`cur_val` before the next test runs.

The limit denominator (`total_abv`) is **fixed** throughout — it is the original
sum of pre-concentration assigned values, mirroring CL!M43 = Portfolio!V61 (a
fixed cell reference in every test's formula).

### 13 Tests Implemented

| # | Test | Qualifying Rule | Limit |
|---|------|----------------|-------|
| 1a | Max 2L & FILO senior lev ≥ 1.50x | 2L OR (FILO AND net_det ≥ 1.50x) | `conc_max_2l_filo_high_lev` |
| 1b | Max Second Lien | loan_type == Second Lien | `conc_max_second_lien` |
| 2 | Max Non-First Lien | 2L OR (FILO AND net_det > 1.00x) | `conc_max_non_first_lien` |
| 3 | Max EBITDA < $5MM | current_ebitda_mm < 5.0 | `conc_max_small_ebitda` |
| 4 | Max Obligors | per-obligor cap, sum excesses | `conc_max_obligor` × total_abv each |
| 5a | Max Largest Industry | largest industry by current value | `conc_max_largest_industry` |
| 5b | Max Second Largest Industry | 2nd largest by current value (post-5a) | `conc_max_second_industry` |
| 5c | Max Other Industries | each remaining industry individually | `conc_max_other_industries` each |
| 6 | Fixed Rate | rate_type == "Fixed" | `conc_fixed_rate` |
| 7 | Max Limited Industry | flags.limited_industry OR industry in policy.limited_industries | `conc_limited_industry` |
| 8 | Max DDTL and Revolver | loan_type in {DDTL, Revolver} | `conc_max_ddtl_revolver` |
| 9a | Max Non-Sponsor | flags.non_sponsor | `conc_max_non_sponsor` |
| 9b | Max Div Recap | flags.div_recap | `conc_max_div_recap` |

All 13 policy percentages come from `PolicyConfig` (loaded from AGENT sheet).
No thresholds are hardcoded.

### Industry ranking

Industries are ranked by their **current values** (post-tests 1–4 haircuts) before
tests 5a/5b/5c run. Industry rankings are re-evaluated after each sub-test (5a then
5b) to avoid stale ranking. Test 5c excludes the top-2 industries identified at
the time of test 5b.

---

## WAAR Calculation

```
raw_waar = total_borrowing_value / total_abv
         = SUMPRODUCT(advance_rates, assigned_value_fractions)
         = mirrors CL!G38
```

**Obligor-count cap** (mirrors Availability!L48):
| Discrete obligors | Cap |
|-------------------|-----|
| < `threshold_mid_diversity` (12) | `cap_low_diversity` |
| ≥ 12 and < `threshold_high_diversity` (20) | `cap_mid_diversity` |
| ≥ 20 | `cap_high_diversity` |

`final_waar = MIN(raw_waar, applicable_cap)`

---

## Availability Calculation

Mirrors Availability sheet L22 / L23 / L25 / L26:

| Test | Formula | Notes |
|------|---------|-------|
| A (L22) | `facility_amount` | $200M constant |
| B (L23) | `net_abv × final_waar + available_collections` | See approximations |
| C (L25) | `total_abv - credit_enhancement_min + available_collections` | See approximations |
| Final (L26) | `MIN(Test A, Test B, Test C)` | |

**Credit enhancement minimum (L41):**
`MAX($15,000,000, sum of 3 largest obligors' pre_conc_values)`

---

## Approximations and Pending Items

| Item | Workbook Cell | This Implementation | Impact |
|------|--------------|---------------------|--------|
| Unfunded exposure equity | L37 (array formula) | Set to $0 | Overstates Tests B and C |
| Interest reserve | L19 | Set to $0 | Overstates Tests B and C |
| Per-obligor top-3 distinct thresholds | Portfolio FE-FG / CL!J54-J56 | Single `conc_max_obligor` cap for all obligors | Minor: top-3 might have different limits |

For the current workbook state, L37 ≈ $0 (no unfunded DDTL/Revolver exposure) and
L19 ≈ $0 (no interest reserve requirement), so the approximation has negligible impact.

---

## Observed Results (2025-02-11 workbook, measurement date 2025-12-31)

| Metric | Value |
|--------|-------|
| Total ABV (eligible pre-conc value) | $120,210,997 |
| Total excess concentration | $29,281,364 |
| Net ABV (after haircuts) | $90,929,633 |
| Raw WAAR (pre-cap) | 60.80% |
| Discrete eligible obligors | 25 |
| WAAR cap applied | 65.00% (≥20 obligors) |
| Final WAAR | 60.80% (unconstrained) |
| Test A (facility) | $200,000,000 |
| Test B (borrowing base) | $60,753,941 ← binding |
| Test C (credit enhancement) | $81,704,750 |
| **Availability** | **$60,753,941** |
| Current advances | $122,300,000 |
| Available cushion | -$61,546,059 (over-advanced) |

**Tests with excess:**
- Max Obligors: $24,911,002 excess (25 individual obligors; limit = $9.0M each)
- Max Largest Industry: $4,370,362 excess (largest industry $28.4M > 20% limit $24.0M)

---

## Testing

### Direct CLI
```bash
cd C:\Users\henry.yan\borrowing-base
PYTHONPATH=src python src\borrowing_base_workbench\calculator.py
```
Prints: eligibility summary, availability block, per-test concentration table,
advance rate distribution.

### Via engine
```python
from borrowing_base_workbench.engine import probe_workbook
r = probe_workbook(r"C:\Users\henry.yan\Downloads\2025-02-11_BDC Borrowing_Base_v8.xlsm")
print(r["metrics"]["availability"])         # 60,753,941
print(r["metrics"]["net_adjusted_bv"])      # 90,929,633
print(r["metrics"]["excess_concentration"]) # 29,281,364
print(r["metrics"]["weighted_avg_advance_rate"])  # 0.608
for t in r["concentration_limits"]:
    if t["excess"] > 0:
        print(t)
```

---

## What Remains for Phase 5 (Scenario Mutation)

1. **Scenario loan construction** — build `LoanRecord + ObligorRecord` from the
   Deal Team Input scenario dict, add to `WorkbookData.loans`, rerun
   `calculate_portfolio()`, return before/after delta
2. **Eligibility for scenario loan** — run same 10 tests on the new synthetic loan
3. **Unfunded equity deduction** — implement L37 array formula if needed
   (currently ~$0 for this portfolio)
4. **Pro forma delta metrics** — availability change, WAAR change, concentration
   impact of the new loan
