# Runtime Architecture — Workbook Import and Session State

## Overview

The app separates workbook **import** (one-time file read) from
**execution** (pure in-memory Python). After import the workbook file
is not accessed again unless the user explicitly reloads it.

```
Workbook file (.xlsm)
      |
      | import_workbook()  — one file read per session
      v
  RuntimeData  ←─────────────────────────────────────┐
      |                                               │
      |── probe_runtime()                             │
      |       └─ calculate_portfolio()               │
      |              └─ PortfolioCalcResult           │
      |                                               │
      └── run_pro_forma_on_runtime(scenario)         │
              │                                       │
              ├─ calculate_portfolio()  (before)      │
              │                                       │
              ├─ build_scenario_records(scenario)     │
              │       └─ LoanRecord + ObligorRecord   │
              │                                       │
              ├─ inject_scenario()                    │
              │       └─ new WorkbookData copy ───────┘
              │          (does NOT mutate RuntimeData)
              │
              └─ calculate_portfolio()  (after)
```

---

## Workbook Import Lifecycle

| Stage | What happens | File required? |
|-------|-------------|----------------|
| `import_workbook(path)` | Opens workbook via openpyxl, reads 6 sheets, closes file | **Yes** |
| Returns `RuntimeData` | All data held in Python objects | No |
| All subsequent calls | Work on `RuntimeData` only | No |
| User clicks "Load Workbook" | App calls `import_workbook()` again | Yes |

The workbook is opened exactly once per load action. `analyze_workbook()` (for
the Admin tab diagnosis) also reads the file, but independently of the data import.
Both calls are batched into the single `_load_workbook()` action in the app.

---

## RuntimeData Lifecycle

```
app startup
    │
    ├─ _try_load_default()
    │       └─ _load_workbook()
    │               ├─ analyze_workbook(path)   → self.diagnosis
    │               └─ import_workbook(path)    → self.runtime_data: RuntimeData
    │
    │   [user changes workbook path and reloads → _load_workbook() again]
    │
    ├─ "Read Current Model" button
    │       └─ probe_runtime(self.runtime_data) → metrics dict
    │
    └─ "Run Pro Forma" button
            └─ run_pro_forma_on_runtime(self.runtime_data, scenario_dict)
                    ├─ computes baseline from runtime_data
                    ├─ injects scenario into a copy
                    └─ computes after-state from the copy
```

`self.runtime_data` is `None` until a workbook has been successfully loaded.
Both `_probe_excel` and `_run_pro_forma` guard against `None` and show an
error message if called before a workbook is loaded.

---

## Scenario Execution Lifecycle

Scenario injection never mutates the base `RuntimeData`:

1. `runtime_data.to_workbook_data()` produces a `WorkbookData` snapshot
2. `inject_scenario(snapshot, loan, obligor, flags)` returns a **new** `WorkbookData`
   with the scenario loan appended — the original snapshot is unchanged
3. Both the before and after portfolios are calculated from their respective
   snapshots
4. After `run_pro_forma_on_runtime` returns, `self.runtime_data` is identical
   to what it was before the call

Multiple scenarios can be run from the same `RuntimeData` baseline in any
order without interference.

---

## What Is File-Backed vs In-Memory

| Data | Backed by | Notes |
|------|-----------|-------|
| Loan tape | In-memory (`RuntimeData.loans`) | Loaded once at import |
| SM Support / obligors | In-memory (`RuntimeData.obligors`) | Loaded once at import |
| VAE records | In-memory (`RuntimeData.vaes`) | Loaded once at import |
| Policy config | In-memory (`RuntimeData.policy`) | Loaded once at import |
| Portfolio flags | In-memory (`RuntimeData.portfolio_flags`) | Loaded once at import |
| Availability metadata | In-memory (`RuntimeData.availability_meta`) | Loaded once at import |
| Workbook structure (Admin tab) | Loaded fresh on each `_load_workbook()` | Used by `analyze_workbook()` only |
| Scenario loans | In-memory only | Never written to any file |
| Calculation results | In-memory only | Never persisted |

---

## Refresh / Reload Path

To pick up workbook changes (new loan tape, updated policy):

1. User clicks **Load Workbook** (or changes the workbook path via Browse)
2. `_load_workbook()` calls both `analyze_workbook()` and `import_workbook()`
3. `self.runtime_data` is replaced with the freshly imported state
4. `self.last_probe_result` is cleared — user must re-run "Read Current Model"

There is no automatic refresh. The session holds the imported state until
explicitly reloaded.

---

## Module Responsibilities

| Module | Responsibility |
|--------|----------------|
| `loader.py` | Reads workbook via openpyxl; owns `RuntimeData` and `WorkbookData` dataclasses; exports `load_workbook()` |
| `engine.py` | Backend API boundary; owns `import_workbook()`, `probe_runtime()`, `run_pro_forma_on_runtime()`; delegates to calculator + scenario |
| `calculator.py` | Pure calculation — eligibility, tiers, WAAR, waterfall, availability; accepts `WorkbookData`; no file I/O |
| `scenario.py` | Builds synthetic records from form dict; `inject_scenario()` returns an immutable copy |
| `app.py` | Tkinter UI; holds `self.runtime_data`; calls engine API only |
