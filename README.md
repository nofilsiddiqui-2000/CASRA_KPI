# CASRA_KPI

Automation for the CASRA Material Master KPI report.

**Business rules and column logic** → see [CASRA KPI.md](CASRA%20KPI.md) (for analysts and process owners).

**Explaining why one part was flagged** → see [Validating KPI Errors.md](Validating%20KPI%20Errors.md) (which file to open, what to search for, what counts as an error).

**This README** → how the code is organized and how to run it.

---

## Code layout (`src/`)

| Module | Role |
|--------|------|
| `Run_KPI.py` | Entry point — run mode prompt, runs pipeline in order |
| `Main_SAP_ZMNM_xl.py` / `Main_SAP_ZMMR2199M_xl.py` | SAP GUI extracts |
| `access_db.py` | KPI build (Access q03–q47 logic) |
| `apply_snp_exceptions.py` | SNP exception pass |
| `generate_hazmat_kpi.py` | HazMat pass. Two modes set by `HAZMAT_MODE`: `skip` (default, no SAP, Hazmat % = 0.00%) and `live` (MM03 validation — not working yet) |
| `generate_kpi_metrics.py` | Dashboard metrics + master append |
| `casra_common.py` | Shared helpers — column lists, dates (`--date-from`/`--date-to`), `config.txt` reader, output paths, SAP GUI session handling, Excel helpers |

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
    └── KPI_Master/
```

`config.txt` in `src/` — SAP credentials and the ZMNM scope filter (`username`, `password`, `zmnm_created_by`). It is **gitignored**: copy `config.example.txt` to `config.txt` and fill it in on each machine that runs the pipeline.

`zmnm_created_by` is the comma-separated list of SAP user IDs the ZMNM extract filters on. Changing it changes which parts are extracted, and therefore every KPI number.

---

### Master metrics file

`KPI_Master/CASRA_KPI_METRICS_MASTER.xlsx` is **appended to**, not rewritten. A run adds one row and touches nothing else, so hand-made changes stay put: corrected numbers, HazMat errors patched in after checking MM03, notes, extra columns, extra sheets and formatting all survive.

The new row is laid out to match the header already in the sheet, so a column you added or moved stays where you put it. If a metric has no column in the sheet, the run says so rather than dropping the value quietly.

The new row also picks up the number format of the cell above it, so a column you
formatted yourself (a percentage, a thousands separator) keeps looking the same.
The three date columns are the exception: they are always written as real dates,
never inheriting an unformatted cell above them, because a date written as a serial
number is what broke the Power BI refreshes.

One caveat: the row goes in below the last row Excel considers used. If stray formatting far down the sheet has inflated that, you will get a gap - delete the empty rows and the next run lands in the right place.

---

## Power BI

The pipeline writes **only to the local `CASRA_KPI_OUTPUT/` folders**. It does not copy anything into SharePoint.

Automatic SharePoint mirroring was removed: it was rewriting the two workbooks Power BI reads, and the date columns were not always written back with the same type, which broke the overnight dashboard refreshes. Publish `SNP_Final/` and `KPI_Master/CASRA_KPI_METRICS_MASTER.xlsx` by hand when you want the dashboards to pick up a new run.

---

## Advanced

**Single step with explicit dates:**

```powershell
python access_db.py --date-from 20260501 --date-to 20260531
python apply_snp_exceptions.py --date-from 20260501 --date-to 20260531
python generate_hazmat_kpi.py --date-from 20260501 --date-to 20260531
python generate_kpi_metrics.py --date-from 20260501 --date-to 20260531
```

### HazMat modes

`HAZMAT_MODE` at the top of `generate_hazmat_kpi.py`:

- `"skip"` (default) — MM03 is never opened. Every HAZ part is treated as correctly
  flagged, so **Hazmat % is always 0.00%** and HazMat errors have to be checked by hand
  in MM03 and patched into `CASRA_KPI_METRICS_MASTER.xlsx`. `Run_KPI.py` prints a reminder
  when this mode is on.
- `"live"` — validates each HAZ part in MM03 (Sales Text across sales orgs 3000 / 4200 /
  1000, plant 3099, distribution channel 00). **Not working yet**: the recorded navigation
  blocks part-way through a run. The MM03 helpers and the remaining steps are at the bottom
  of `generate_hazmat_kpi.py`.

`ZMNM_FILE` at the top of the same file points the step at a specific ZMNM workbook when
debugging; leave it blank for a normal pipeline run.

Live mode drives **MM03** through SAP GUI scripting, so it needs an SAP logon (`config.txt`
credentials, same `PR2` / client `320` as the extracts) and SAP GUI scripting enabled.

### Data Quality reminder

`Run_KPI.py` asks you to confirm the Data Quality workbook before the first step runs. It
shows the expected path and when the file was last modified, and warns if it predates the
period you selected. The filename lives at the top of `apply_snp_exceptions.py`
(`DATAQUALITY_FILE`) and has to be updated when you download a new month.
