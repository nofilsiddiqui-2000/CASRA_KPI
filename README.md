# CASRA_KPI

Automation for the CASRA Material Master KPI report.

**Business rules and column logic** → see [CASRA KPI.md](CASRA%20KPI.md) (for analysts and process owners).

**This README** → how the code is organized and how to run it.

---

## Code layout (`src/`)

| Module | Role |
|--------|------|
| `Run_KPI.py` | Entry point — run mode prompt, runs pipeline in order |
| `Main_SAP_ZMNM_xl.py` / `Main_SAP_ZMMR2199M_xl.py` | SAP GUI extracts |
| `access_db.py` | KPI build (Access q03–q47 logic) |
| `apply_snp_exceptions.py` | SNP exception pass |
| `generate_hazmat_kpi.py` | HazMat pass — validates ZMNM HAZ parts in MM03 (Sales Text across 3 sales orgs) and appends them with Check_Hazards 0/1. Needs SAP GUI |
| `generate_kpi_metrics.py` | Dashboard metrics + master append |
| `kpi-metrics-manual.py` | On-demand metrics from user-supplied ZMNM/ZMMR/DQ files (no HazMat pass; no KPI Master update) |
| `casra_common.py` | Shared helpers — column lists, dates (`--date-from`/`--date-to`), `config.txt` reader, output/SharePoint paths, Excel helpers |

---

## How to run

```powershell
cd D:\Bombardier\CASRA_KPI\src
python Run_KPI.py
```

```
CASRA KPI - Run Configuration
----------------------------------------
  1) Automated  -  Previous calendar month
  2) Manual     -  Enter a custom date range

Select run mode [1/2]:
```

### Option 1 — Automated

Press `1`. Uses the previous calendar month for all steps (example: run on 1 Jun 2026 → `20260501` to `20260531`).

### Option 2 — Manual

Press `2`, then enter dates as **YYYYMMDD**:

```
  Start date (YYYYMMDD): 20260330
  End date   (YYYYMMDD): 20260503
```

### Pipeline order

1. `Main_SAP_ZMNM_xl.py`
2. `Main_SAP_ZMMR2199M_xl.py`
3. `access_db.py`
4. `apply_snp_exceptions.py`
5. `generate_hazmat_kpi.py`
6. `generate_kpi_metrics.py`

---

## Output folders

Under `CASRA_KPI_OUTPUT/`:

| Folder | File |
|--------|------|
| `Intermediate/` | `CASRA_KPI_OUTPUT_<date_from>_<date_to>.xlsx` |
| `SNP_Final/` | `CASRA_KPI_OUTPUT_<date_from>_<date_to>_FINAL.xlsx` |
| `HazMat_KPI/` | `CASRA_HAZMAT_KPI_<run date>.xlsx` (HAZ parts, MM03 Sales-Text results, skipped parts, summary) |
| `KPI_Metrics/` | `CASRA_KPI_METRICS_<date_from>_<date_to>.xlsx` |
| `KPI_Master/` | `CASRA_KPI_METRICS_MASTER.xlsx` |
| `KPI_Metrics_Manual/` | `CASRA_KPI_METRICS_MANUAL_<today>.xlsx` (manual script only) |

`<date_from>` and `<date_to>` = period start and end from your run.

---

## Folder layout (project root)

```
CASRA-KPI-AUTOMATION/
├── SAP_Extracts/
│   ├── ZMNM/
│   └── ZMMR2199M/
├── LookUp Tables/
│   ├── RuleSloc.xlsx
│   ├── QMatMissing.xlsx
│   ├── QMATRules.xlsx
│   └── Data_Quality_*.xlsx          (monthly download; set filename in apply_snp_exceptions.py)
└── CASRA_KPI_OUTPUT/
    ├── Intermediate/
    ├── SNP_Final/
    ├── HazMat_KPI/
    ├── KPI_Metrics/
    ├── KPI_Master/
    └── KPI_Metrics_Manual/
```

`config.txt` in `src/` — SAP credentials (`username`, `password`, `asset_num`).

---

## Power BI (SharePoint sync)

After each run, two files are copied to your SharePoint-synced folder:

| Subfolder | File |
|-----------|------|
| `SNP_Final/` | `CASRA_KPI_OUTPUT_<date_from>_<date_to>_FINAL.xlsx` (updated after HazMat pass) |
| `KPI_Master/` | `CASRA_KPI_METRICS_MASTER.xlsx` |

Update the root path in `src/casra_common.py`:

```python
SHAREPOINT_SYNC_ROOT = Path(r"C:\Users\B1020000\Bombardier\SharePoint-Sync\CASRA_KPI_PowerBI")
```

Create `SNP_Final` and `KPI_Master` inside that folder once, then connect Power BI there.

---

## Advanced

**Single step with explicit dates:**

```powershell
python access_db.py --date-from 20260501 --date-to 20260531
python apply_snp_exceptions.py --date-from 20260501 --date-to 20260531
python generate_hazmat_kpi.py --date-from 20260501 --date-to 20260531
python generate_kpi_metrics.py --date-from 20260501 --date-to 20260531
```

Optional, at the top of `generate_hazmat_kpi.py` (leave blank/empty for the normal pipeline path):
- `ZMNM_FILE` — point at a specific ZMNM workbook when debugging.
- `MATERIALS_OVERRIDE` — a fixed list of part numbers to validate in MM03 instead of the HAZ parts read from ZMNM (handy for testing the SAP path).

The HazMat step drives **MM03** through SAP GUI scripting, so it needs an SAP logon (`config.txt` credentials, same `PR2` / client `320` as the extracts) and SAP GUI scripting enabled.

**Manual metrics only** (does not update KPI Master; Hazmat % = 0 unless you run the HazMat step separately):

```powershell
python kpi-metrics-manual.py
```

Fill in `ZMNM_FILE`, `ZMMR_FILE`, and `DATA_QUALITY_FILE` at the top of `kpi-metrics-manual.py` before running.
