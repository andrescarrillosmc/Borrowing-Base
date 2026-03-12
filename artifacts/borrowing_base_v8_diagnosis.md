# Borrowing Base v8 Diagnosis

- Workbook: C:\Users\Andres.Carrillo\OneDrive - Star Mountain Capital\03_Operations\01_Financial\Borrowing_Base\Product\2025-02-11_BDC Borrowing_Base_v8.xlsx
- Version: v2.1
- Measurement date: 2025-12-31 00:00:00
- Borrower: Star Mountain Lower Middle-Market Capital Corp.
- Inputs mapped: 36
- Outputs mapped: 24
- Validation rules: 32
- Write policies: 18

## Monthly Update Tabs

- Loan Tape - Settled
- SM Support
- Availability
- Deal Team Input
- Memory
- MAPPING_LAYER
- Portfolio

## Structural Risks

- Cached output values are mostly blank, so exact numbers need an Excel-backed calculation engine.
- Portfolio formulas use fixed Loan Tape ranges, which creates a monthly range-extension obligation.
- Deal Team Input is the UX layer, but MAPPING_LAYER is the real operational contract and should drive the app.
- AGENT and Availability should be admin-governed because they encode legal or facility-wide logic.

## Recommended App Shape

- Use the workbook as the governed monthly logic template while the app owns workflow, validation, audit logging, and side-by-side presentation.
- Drive form fields, outputs, and write permissions from MAPPING_LAYER instead of hard-coding them in the app.
- Split the product into Deal Team mode and Admin mode. Deal Team runs scenarios; Admin manages workbook versions and monthly health checks.
- Treat Loan Tape - Settled, SM Support, Availability metadata, and Portfolio range health as the main monthly maintenance surfaces.

## Capacity Diagnostics

- Loan Tape within current Portfolio source ranges: last row 45, modeled ceiling 55, remaining headroom 10. Portfolio formulas reference fixed Loan Tape ranges ending at row 55.
- SM Support append zone before totals: last row 159, modeled ceiling 63, remaining headroom -96. Totals begin at row 64, so scenario rows should be inserted before that summary block.
- Portfolio insertion marker: last row 60, modeled ceiling n/a, remaining headroom n/a. The sheet contains a manual instruction row showing where new collateral rows were expected to be added.