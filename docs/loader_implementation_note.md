# Workbook Data Loader — Implementation Note

**Branch:** `engine-detach-v1`
**Phase:** 2 of 6 (Workbook Data Loader)

---

## What Was Built

### New file: `src/borrowing_base_workbench/loader.py`

A single-module workbook reader.  Opens the `.xlsm` file with `openpyxl`
(`data_only=True`) and returns structured Python dataclasses.  No Excel COM,
no PowerShell, no recalculation.

Public entry point:
```python
from borrowing_base_workbench.loader import load_workbook_data

data = load_workbook_data(Path("path/to/workbook.xlsm"))
# returns WorkbookData
```

---

## Workbook Tabs Loaded

| Tab | Rows Loaded | What It Represents | Classification |
|-----|-------------|-------------------|----------------|
| **Loan Tape - Settled** | 5–55 (41 active) | One row per loan in the settled portfolio | RAW INPUT |
| **SM Support** | 3–64 (45 active) | One row per obligor: EBITDA, leverage, FMV, balances | RAW INPUT + formula proxy |
| **VAE** | 5–200 (31 active) | Value Adjustment Events: per-borrower credit deterioration flags | RAW INPUT |
| **AGENT** | Key cells only | All facility thresholds, advance rates, concentration limits | POLICY CONFIG |
| **Concentration Limits** | E81:E153 (73 entries) | Industry taxonomy (S&P classifications) | POLICY CONFIG |
| **Portfolio** | Cols BV/BW/BX/CA/CQ/CW/CX only | Manual per-loan boolean flags (non-sponsor, div-recap, limited-industry, DIP, agent overrides) | MANUAL FLAG |
| **Availability** | L16, L18, L51, F8 | Facility amount, cash, current advances, measurement date | RAW INPUT (constants) |

**Not loaded** (computed by calculator.py in Phase 3):
- Portfolio eligibility columns (E–N, P)
- Portfolio collateral tier (AF), assigned value (AL)
- Portfolio concentration test columns (EI–HD)
- All Availability formula outputs (L22–L26, L33–L35, L41–L48)
- Concentration Limits advance rate and excess tables

---

## Loaded Data Structures

### `LoanRecord` — one per row in Loan Tape - Settled
Key fields: `obligor_name`, `loan_type`, `denomination`, `olb`, `fmv`,
`purchase_price`, `industry`, `cov_lite`, `pik_pct`, `inclusion_date`,
`initial_total_leverage`, `net_detachment`, `net_attachment`

**Formula-proxy columns** (computed by loader from SM Support join, not from cached Excel values):
- `commitment_balance`, `olb` — from `ObligorRecord.debt_balance`
- `fmv` — from `ObligorRecord.sm_fair_value`
- `current_ebitda_mm` — from `ObligorRecord.ltm_adj_ebitda / 1,000,000`
- `net_detachment`, `net_attachment` — from `ObligorRecord.net_detachment/net_attachment`

**EBITDA scaling note:** `initial_ebitda_mm` (col AV) is stored in the workbook
in millions (e.g., `6.444` = $6,444,000).  The loader multiplies by 1,000,000.
Current EBITDA comes from SM Support in full dollars.

**Inclusion date note:** `inclusion_date` (col AO) is loaded separately from
`investment_date` (col U).  The Portfolio maturity test uses `(maturity - inclusion_date) / 365`,
not investment date.  In the current workbook AO ≠ U for many loans.

### `ObligorRecord` — one per row in SM Support
Key fields: `obligor_name`, `debt_balance`, `sm_fair_value`, `ltm_adj_ebitda`,
`drawn_revolver`, `first_out`, `total_sm_balance`, `cash_balance`,
`net_detachment`, `net_attachment`, `interest_coverage`

`net_detachment` = `(revolver + first_out - cash) / ebitda` — computed by loader, not cached.
`net_attachment` = `(revolver + first_out + pari + sm - cash) / ebitda` — same.
`sm_fair_value` = `debt_balance × mark_avg` where `mark_avg = (mark_low + mark_high) / 2`.

### `VaeRecord` — one per row in VAE sheet
Key fields: `borrower` (join key), `event_type`, `vae_date`,
`ebitda_at_vae`, `vae_agent_assigned_value`

The operative field for the engine is `vae_agent_assigned_value` (col P).
Engine will use `min(fmv, vae_agent_assigned_value, collateral_tier_value)` for assigned value.

### `PolicyConfig` — from AGENT sheet
Contains: all 7 advance rates, 3 portfolio caps with obligor-count thresholds,
4 collateral tiers, 13 concentration limit percentages, 4 limited industries,
73-entry industry taxonomy.

Collateral tiers are loaded as a sorted list of `CollateralTier(max_lev, applicable_pct)`:
```
Tier 1: FL ≤ 3.25x / FILO-2L ≤ 3.75x → 100%
Tier 2: FL ≤ 4.00x / FILO-2L ≤ 4.25x → 92.5%
Tier 3: FL ≤ 5.00x / FILO-2L ≤ 5.00x → 85%
Tier 4: > 5.00x → 0%  (sentinel appended by loader)
```

### `ManualPortfolioFlags` — from Portfolio sheet (per obligor)
Six boolean flags with no upstream formula source:

| Flag | Portfolio Col | Affects |
|------|--------------|---------|
| `limited_industry` | BV | Concentration Test 7 (10% cap) |
| `non_sponsor` | BW | Concentration Test 9a (15% cap) |
| `div_recap` | BX | Concentration Test 9b (10% cap) |
| `is_dip` | CA | Eligibility Test 2 (PIK/DIP) |
| `agent_addback_discretion` | CQ | Permitted EBITDA calculation |
| `agent_post_inclusion_haircut` | CW | Permitted EBITDA calculation |

For scenario loans (not in workbook), all flags default `False`.

### `AvailabilityMeta` — three raw cells from Availability
- `facility_amount` = L16 = $200,000,000
- `available_collections` = L18 (cash, currently $5,470,978)
- `current_advances` = L51 = $122,300,000 (hardcoded constant)
- `measurement_date` = F8 = 2025-12-31

### `WorkbookData` — top-level container
Returned by `load_workbook_data()`.  Provides helper methods:
- `obligor_by_name(name)` — lookup `ObligorRecord`
- `vaes_for(obligor_name)` — all `VaeRecord` for a borrower
- `flags_for(obligor_name)` — `ManualPortfolioFlags` or safe default
- `summary()` — dict for logging/debug

---

## What Remains for Phase 3 (calculator.py)

The loader gives the engine clean, typed data.  Calculator.py will use it to implement:

1. **Eligibility** — 10 boolean tests per loan → `eligible: bool`
2. **Collateral tier** — leverage lookup in `PolicyConfig.collateral_tiers`
3. **VAE override** — `min(fmv, vae_assigned, tier_value)` → `assigned_value`
4. **Advance rate bucket** — EBITDA and leverage-based → `rate: float`
5. **Concentration tests** — 13 sequential haircut tests → per-loan excess
6. **Weighted average advance rate** — `SUMPRODUCT` + obligor-count cap
7. **Availability** — `MIN(test_a, test_b, test_c)`
8. **Scenario loan construction** — build `LoanRecord + ObligorRecord` from scenario dict

---

## Testing the Loader

### Direct CLI test
```bash
cd C:\Users\henry.yan\borrowing-base
python src\borrowing_base_workbench\loader.py
```
Prints: loan/obligor/VAE counts, sample rows, policy values, portfolio flags.

### Via engine
```python
from borrowing_base_workbench.engine import probe_workbook

result = probe_workbook(r"C:\Users\henry.yan\Downloads\2025-02-11_BDC Borrowing_Base_v8.xlsm")
print(result["status"])            # "ok"
print(result["loader_summary"])    # counts and metadata
print(result["metrics"])           # zeroed until Phase 3
print(result["concentration_limits"])  # 13 rows with live policy percentages
```

### Expected output
```
loan_count: 41
obligor_count: 45
vae_count: 31
facility_amount: $200,000,000
current_advances: $122,300,000
measurement_date: 2025-12-31
industry_taxonomy_count: 73
```

### Via app
Run `run_app.bat`, load the workbook, click **Read Current Model**.
- The Results tab shows all-zero metrics (expected — Phase 3 not yet implemented)
- The concentration limits table shows the correct 13 test names and percentages
- `current_advances` populates correctly ($122,300,000)
- No errors or crashes
