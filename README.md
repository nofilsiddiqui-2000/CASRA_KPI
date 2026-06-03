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
| `access-db.py` | KPI build (Access q03–q47 logic) |
| `apply_snp_exceptions.py` | SNP exception pass |
| `generate_kpi_metrics.py` | Dashboard metrics + master append |
| `kpi-metrics-manual.py` | On-demand metrics from a chosen FINAL file |
| `casra_paths.py` | Local output folders + SharePoint sync paths |
| `casra_dates.py` | `--date-from` / `--date-to` |
| `casra_constants.py` | Shared column lists |
| `casra_excel.py` | Shared Excel helpers |
| `casra_config.py` | `config.txt` reader (SAP) |

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
3. `access-db.py`
4. `apply_snp_exceptions.py`
5. `generate_kpi_metrics.py`

---

## Output folders

Under `CASRA_KPI_OUTPUT/`:

| Folder | File |
|--------|------|
| `Intermediate/` | `CASRA_KPI_OUTPUT_<date_from>.xlsx` |
| `SNP_Final/` | `CASRA_KPI_OUTPUT_<date_from>_FINAL.xlsx` |
| `KPI_Metrics/` | `CASRA_KPI_METRICS_<date_from>.xlsx` |
| `KPI_Master/` | `CASRA_KPI_METRICS_MASTER.xlsx` |
| `KPI_Metrics_Manual/` | `CASRA_KPI_METRICS_MANUAL_<today>.xlsx` (manual script only) |

`<date_from>` = period start date from your run.

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
| `SNP_Final/` | `CASRA_KPI_OUTPUT_<date_from>_FINAL.xlsx` |
| `KPI_Master/` | `CASRA_KPI_METRICS_MASTER.xlsx` |

Update the root path in `src/casra_paths.py`:

```python
SHAREPOINT_SYNC_ROOT = Path(r"C:\Users\B1020000\Bombardier\SharePoint-Sync\CASRA_KPI_PowerBI")
```

Create `SNP_Final` and `KPI_Master` inside that folder once, then connect Power BI there.

---

## Advanced

**Single step with explicit dates:**

```powershell
python access-db.py --date-from 20260501 --date-to 20260531
python apply_snp_exceptions.py --date-from 20260501 --date-to 20260531
```

**Manual metrics only** (does not update KPI Master):

```powershell
python kpi-metrics-manual.py
python kpi-metrics-manual.py --input "C:\path\to\SNP_Final\CASRA_KPI_OUTPUT_20260501_FINAL.xlsx"
```
