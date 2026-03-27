# Backend Seam — Implementation Note

**Branch:** `detached`
**Phase:** 1 of 6 (Backend Abstraction Layer)

---

## What Changed

### New file: `src/borrowing_base_workbench/engine.py`

The Python execution boundary between the Tkinter frontend and the calculation backend. Exposes two public functions:

```python
probe_workbook(workbook_path: str | Path) -> dict
run_pro_forma(workbook_path: str | Path, scenario: dict) -> dict
```

Both return dicts with the same shape as the retired PowerShell scripts. Current implementation is **stubbed**: all numeric outputs are `0.0`, eligibility is `"Yes"`, and concentration limits are zeroed. This is intentional — it keeps the UI fully renderable while the real engine is built in the next phase.

### Modified: `src/borrowing_base_workbench/app.py`

**Line 12 — import swap (the only structural change):**
```python
# Before
from borrowing_base_workbench.excel_runner import probe_excel_workbook, run_pro_forma_workbook

# After
from borrowing_base_workbench.engine import probe_workbook as probe_excel_workbook, run_pro_forma as run_pro_forma_workbook
```

**`_run_pro_forma()` — removed `script_path`:**
```python
# Before
script_path = Path(__file__).with_name("run_pro_forma.ps1")
result = run_pro_forma_workbook(self.workbook_path.get(), script_path, values)

# After
result = run_pro_forma_workbook(self.workbook_path.get(), values)
```

**`_probe_excel()` — removed `script_path`:**
```python
# Before
script_path = Path(__file__).with_name("excel_probe.ps1")
result = probe_excel_workbook(self.workbook_path.get(), script_path)

# After
result = probe_excel_workbook(self.workbook_path.get())
```

### Modified: `src/borrowing_base_workbench/analysis.py`

Replaced the `subprocess` / PowerShell `Copy-Item` call in `_stage_readable_copy()` with `shutil.copy2`. Removed the `subprocess` import. The `analysis.py` module now has zero subprocess or PowerShell dependency.

---

## What Still Depends on Excel

| File | Dependency | Status |
|------|-----------|--------|
| `excel_runner.py` | Wraps PowerShell via `subprocess` | **Unused** — retained for reference, will be deleted in Phase 6 |
| `excel_probe.ps1` | Excel COM automation for baseline read | **Unused** — retained for reference |
| `run_pro_forma.ps1` | Excel COM automation for scenario write | **Unused** — retained for reference |
| `engine.py` | **None** — stub only | ✅ No Excel dependency |
| `analysis.py` | **None** — only uses `openpyxl` + `shutil` | ✅ No Excel dependency |

The app no longer calls any PowerShell or Excel COM code at runtime. The three legacy files remain on disk but are not imported anywhere.

---

## How to Test the Seam

### 1. Verify the app launches
```
run_app.bat
```
The app should open normally. No error on startup.

### 2. Verify "Read Current Model" works
- Load the workbook from the Admin tab
- Click **Read Current Model** on the Scenario tab
- Expected: The Results tab populates with zeroed values (availability = $0, all concentration tests show $0)
- Expected: No error dialog, no crash

### 3. Verify "Run Pro Forma" works
- Fill in the required scenario fields
- Click **Run Pro Forma**
- Expected: Results tab shows before/after comparison with zeroed values
- Expected: Eligibility shows `"Yes"`, no failed tests
- Expected: Commentary appears

### 4. Confirm no PowerShell is invoked
Run the app with PowerShell execution disabled or with `ExecutionPolicy Restricted` and confirm neither "Read Current Model" nor "Run Pro Forma" fails with a PowerShell error.

---

## Next Phase

**Phase 2: Workbook Data Loader**

Create `loader.py` to load live data from the workbook via `openpyxl`:
- `load_loan_tape()` → `list[LoanRecord]`
- `load_sm_support()` → `list[ObligorRecord]`
- `load_vae()` → `list[VAERecord]`
- `load_policy()` → `PolicyConfig`
- `load_manual_flags()` → `dict[str, PortfolioFlags]` (Portfolio cols BV, BW, BX, CQ, CW, CX)

Once the loader exists, `engine.probe_workbook()` and `engine.run_pro_forma()` stubs are replaced with real calls to `loader.load_workbook()` and `calculator.run()`.
