# Borrowing Base Workbench

Desktop application for Star Mountain Capital's BDC borrowing base model.
Reads the Excel workbook as a data source and computes all portfolio metrics
in pure Python — no Excel runtime, no COM automation, no PowerShell.

## What it does

- **Baseline probe** — reads the current workbook state and computes availability,
  WAAR, concentration waterfall, and per-asset eligibility
- **Pro forma scenario** — adds a synthetic loan to an in-memory copy of the
  portfolio and shows before/after delta across all metrics
- **Pre-flight validation** — checks workbook health and surfaces commentary
  before running calculations
- **Workbook diagnostics** — parses `MAPPING_LAYER` into structured metadata

## Architecture

```
app.py  (Tkinter UI)
  └── engine.py  (backend boundary)
        ├── loader.py      reads workbook via openpyxl — no Excel required
        ├── calculator.py  eligibility tests, advance rates, concentration waterfall, WAAR
        └── scenario.py    constructs synthetic LoanRecord/ObligorRecord for pro forma runs
```

See [docs/architecture.md](docs/architecture.md) for a detailed walkthrough.

## Run

```powershell
python .\src\borrowing_base_workbench\app.py
```

Or:

```powershell
.\run_app.bat
```

The workbook path is set in the Admin tab on first run.

## Tests

```powershell
# From the repo root
cd C:\Users\henry.yan\borrowing-base
PYTHONPATH=src python tests/test_engine.py
```

96 tests covering eligibility logic, collateral tier bucketing, advance rate
selection, VAE handling, scenario injection, and end-to-end integration.
Integration tests skip automatically if the workbook file is absent.

## Parity validation

```powershell
PYTHONPATH=src python tests/validate_parity.py
```

Compares all engine outputs against known-good reference values.
See [parity_report.md](parity_report.md) for results.

## Workbook

The app reads (never writes) the Excel workbook:

`2025-02-11_BDC Borrowing_Base_v8.xlsm`

Sheets used: `Loan Tape - Settled`, `SM Support`, `AGENT`, `VAE Log`,
`CL` (concentration limits), `Availability`.

## Docs

| File | Contents |
|------|----------|
| [docs/architecture.md](docs/architecture.md) | System design, module responsibilities, data flow |
| [parity_report.md](parity_report.md) | Validation results — engine vs reference values |
| [docs/merge_prep.md](docs/merge_prep.md) | Branch change summary and merge recommendation |
