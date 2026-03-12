# Borrowing Base Workbench

Desktop scaffold for turning `Borrowing_Base_v8` into a standalone operator app.

## What it does now

- Diagnoses the workbook structure from the live spreadsheet.
- Parses the `MAPPING_LAYER` into app-ready input, output, validation, and write-policy metadata.
- Surfaces the monthly operator checklist for the tabs most likely to change.
- Provides a deal-team scenario form driven by the workbook's own mapping sheet.
- Runs pre-flight validation and commentary for a pro forma scenario.
- Probes Excel automation so the app can eventually use the workbook itself as the calculation engine.

## Current constraint

The app now targets the trusted workbook saved in the product folder:

`C:\Users\Andres.Carrillo\OneDrive - Star Mountain Capital\03_Operations\01_Financial\Borrowing_Base\Product\2025-02-11_BDC Borrowing_Base_v8.xlsx`

Current implication:

- The analyzer, admin panel, scenario form, and live-model read path are working.
- The app still does not write scenario rows back into workbook copies, so the full governed pro forma execution path is not enabled yet.
- `Read Current Model` is read-only. `Clear Scenario` only resets the UI.

## Run

From this folder:

```powershell
python .\src\borrowing_base_workbench\app.py
```

Or use:

```powershell
.\run_app.bat
```
