# Branch Comparison — Pre-Merge Analysis

**Date:** 2026-03-27

---

## Branch Topology

```
bd0adc4  test commit
c88b60c  Delete README.md
843daec  Initial commit: BorrowingBaseWorkbench
f403317  Merge branch 'main' of ...         <-- HEAD of main
          |
          +-- 2c0fb4b  Phase 1: Introduce Python backend seam  <-- HEAD of detached
                |
                +-- 0161c3a  Phase 2: workbook data loader
                +-- 1e0efb8  Phase 3: asset-level calculation engine
                +-- 7b2c807  Phase 4: concentration waterfall + availability
                +-- a1bb5d2  Phase 5: pro forma scenario execution
                +-- 7a55097  Phase 6: parity validation + regression tests
                +-- 689a442  Phase 7: cleanup, documentation, merge prep    <-- HEAD of engine-detach-v1
```

---

## Commit Containment

| Branch | Commits ahead of main | Commits NOT in engine-detach-v1 |
|--------|----------------------|----------------------------------|
| `main` | — | — |
| `detached` | 1 | **0** |
| `engine-detach-v1` | 7 | — |

**`detached` is fully superseded.**

`git log origin/detached ^origin/engine-detach-v1` returns nothing — every commit
on `detached` is already in `engine-detach-v1`.

---

## What engine-detach-v1 Adds to main

### New files
| File | Description |
|------|-------------|
| `src/.../loader.py` | Workbook reader via openpyxl (772 lines) |
| `src/.../calculator.py` | Eligibility, tiers, WAAR, concentration waterfall (1000 lines) |
| `src/.../scenario.py` | Pro forma record construction (360 lines) |
| `src/.../engine.py` | Backend API — `probe_workbook`, `run_pro_forma` (272 lines) |
| `tests/test_engine.py` | 96 regression tests (1085 lines) |
| `tests/validate_parity.py` | Parity validation script (439 lines) |
| `parity_report.md` | Validation results |
| `docs/architecture.md` | System design document |
| `docs/merge_prep.md` | Merge preparation note |
| `docs/` | Phase implementation notes (moved from root) |

### Modified files
| File | Change |
|------|--------|
| `src/.../app.py` | Import swap + removed `script_path` args at 2 call sites |
| `src/.../engine.py` | Fully implemented (was stub on `main`) |
| `src/.../analysis.py` | Minor updates |
| `README.md` | Full rewrite |

### Deleted files
| File | Reason |
|------|--------|
| `src/.../excel_runner.py` | Dead — replaced by `engine.py` |
| `src/.../excel_probe.ps1` | Dead — Excel COM automation no longer used |
| `src/.../run_pro_forma.ps1` | Dead — Excel COM automation no longer used |

**Total delta:** +5,303 / −471 lines across 21 files.

---

## Recommendation: Can detached Be Deleted?

**Yes.** `detached` contains exactly one commit beyond `main` (Phase 1 backend seam),
and that commit is the base of `engine-detach-v1`. Deleting `detached` loses nothing;
the full history is preserved in `engine-detach-v1`.

---

## Recommended Merge Strategy for engine-detach-v1 → main

**Squash merge** (recommended).

The branch has 7 incremental phase commits plus test fixup commits. The full change
is coherent and self-contained. A single squash commit gives `main` a clean, readable
history without noise from intermediate phases.

### Exact commands

```powershell
# Step 1 — archive current main state (run BEFORE the merge)
git checkout main
git pull origin main
git branch archive/main-pre-engine-detach main
git push origin archive/main-pre-engine-detach

# Step 2 — squash merge
git merge --squash engine-detach-v1

# Step 3 — commit (edit the message if desired)
git commit -m "Replace Excel/PowerShell backend with pure Python calculation engine

- loader.py: reads workbook via openpyxl, no Excel runtime required
- calculator.py: eligibility tests, concentration waterfall, WAAR, availability
- scenario.py: in-memory pro forma scenario injection
- engine.py: backend API seam (probe_workbook, run_pro_forma)
- app.py: one-line import swap, no UI changes
- Removed: excel_runner.py, excel_probe.ps1, run_pro_forma.ps1
- Added: 96 regression tests, parity validation, architecture docs

All 17 reference metrics match exactly. 96/96 tests pass.
No Excel runtime dependency remains."

# Step 4 — push
git push origin main

# Step 5 — delete superseded branches (optional, after confirming main is good)
git push origin --delete detached
git push origin --delete engine-detach-v1
git branch -d detached
git branch -d engine-detach-v1
```

### Alternative: merge commit (non-squash)

If full phase-by-phase history on `main` is preferred:

```powershell
git checkout main
git merge --no-ff engine-detach-v1 -m "Merge engine-detach-v1: replace Excel backend with pure Python engine"
git push origin main
```

This preserves all 7 commits but creates a merge bubble on main's linear history.

---

## Pre-Merge Checklist

- [x] `engine-detach-v1` is up-to-date with origin
- [x] All 96 regression tests pass
- [x] Parity validation passes (17/17 reference checks)
- [x] `detached` is fully contained in `engine-detach-v1`
- [x] Dead code removed (excel_runner, ps1 scripts)
- [x] README and docs updated
- [ ] Archive branch created from current `main`
- [ ] Reviewer sign-off
- [ ] Squash merge executed
