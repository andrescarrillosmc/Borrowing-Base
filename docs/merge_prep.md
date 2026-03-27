# Merge Preparation — engine-detach-v1

**Branch:** `engine-detach-v1`
**Base:** `main`
**Date:** 2026-03-27

---

## What Changed from main

### New files

| File | Description |
|------|-------------|
| `src/.../engine.py` | Public backend API — replaces PowerShell bridge |
| `src/.../loader.py` | Workbook reader via openpyxl |
| `src/.../calculator.py` | Portfolio calculation engine (eligibility, tiers, WAAR, waterfall) |
| `src/.../scenario.py` | Pro forma scenario record construction |
| `tests/test_engine.py` | 96 regression tests |
| `tests/validate_parity.py` | Parity validation script |
| `parity_report.md` | Validation results |
| `docs/architecture.md` | System design document |
| `docs/merge_prep.md` | This file |
| `docs/` | Phase implementation notes (moved from root) |

### Modified files

| File | Change |
|------|--------|
| `src/.../app.py` | One-line import swap + removed `script_path` arguments from two call sites |
| `src/.../engine.py` | Docstrings cleaned (removed stale PowerShell references) |
| `README.md` | Full rewrite — describes current Python-only architecture |

### Deleted files

| File | Reason |
|------|--------|
| `src/.../excel_runner.py` | Replaced by engine.py; no longer imported |
| `src/.../excel_probe.ps1` | Excel COM script; replaced by pure Python |
| `src/.../run_pro_forma.ps1` | Excel COM script; replaced by pure Python |

---

## Legacy Code Remaining

None. All Excel / PowerShell execution paths have been removed from the active
codebase. The retired files existed only in this branch and are now deleted.

The `analysis.py` module uses `shutil.copy2` for workbook staging (not PowerShell),
which is unchanged.

---

## Unchanged from main

- `app.py` UI logic and layout — no visual or behavioral changes to the frontend
- `validation.py` — pre-flight commentary unchanged
- `analysis.py` — workbook structure diagnosis unchanged
- All test fixtures in `tests/test_analysis.py`

---

## Merge Strategy

**Recommended: squash merge into main**

Reason: the branch contains 6 incremental phase commits and several fixup
commits from test development. The cumulative change is coherent and
self-contained — a single merge commit gives a clean history on main.

Suggested merge commit message:
```
Replace Excel/PowerShell backend with pure Python calculation engine

- New: loader.py — reads workbook via openpyxl, no Excel runtime required
- New: calculator.py — eligibility, concentration waterfall, WAAR, availability
- New: scenario.py — in-memory pro forma scenario injection
- New: engine.py — backend API seam (probe_workbook, run_pro_forma)
- Updated: app.py — one-line import swap, no UI changes
- Removed: excel_runner.py, excel_probe.ps1, run_pro_forma.ps1
- Added: 96 regression tests, parity validation, architecture docs

All 17 reference metrics match exactly. 96/96 tests pass.
No Excel runtime dependency remains.
```

---

## Pre-merge Checklist

- [x] 96/96 regression tests pass (`python tests/test_engine.py`)
- [x] Parity validation passes on all 17 reference metrics (`python tests/validate_parity.py`)
- [x] `app.py` imports only from `engine.py` — no PowerShell paths
- [x] Dead code deleted (`excel_runner.py`, `*.ps1`)
- [x] README updated
- [x] Architecture document written
- [x] Branch is ahead of origin on `engine-detach-v1`
- [ ] Reviewer sign-off
- [ ] Merge to main
