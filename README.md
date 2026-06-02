# CASRA_KPI

Automation for the CASRA Material Master KPI report.

The pipeline:

1. Pulls two SAP extracts (`ZMNM` and `ZMMR2199M`) via SAP GUI scripting.
2. Runs the access-db logic (a Python rewrite of the old MS Access queries) to flag data-quality errors and produce the intermediate KPI output.
3. Applies SNP exceptions from the manually-downloaded Data Quality file to clear out parts that are known exceptions.
4. Writes the final KPI Excel file used for reporting and visuals.

Everything is driven from a single entry point: `src/Run_KPI.py`.

---

## How to run

Open a terminal, go to the project's `src` folder, and run:

```powershell
cd D:\Bombardier\CASRA_KPI\src
python Run_KPI.py
```

You'll be prompted to pick a run mode:

```
CASRA KPI - Run Configuration
----------------------------------------
  1) Automated  -  Previous calendar month
  2) Manual     -  Enter a custom date range

Select run mode [1/2]:
```

### Option 1 — Automated (previous month)

Just press `1` and hit Enter. The pipeline figures out the previous calendar month on its own and runs end-to-end. No further input required.

Example — running on **June 1, 2026**:

```
Select run mode [1/2]: 1

  Using previous month: 20260501 -> 20260531
```

It will then run, in order:
- `Main_SAP_ZMNM_xl.py`
- `Main_SAP_ZMMR2199M_xl.py`
- `access-db.py`
- `apply_snp_exceptions.py`

### Option 2 — Manual (custom date range)

Press `2` and hit Enter, then type the start and end dates in **YYYYMMDD** format when prompted.

Example — running for **March 30, 2026 to May 3, 2026**:

```
Select run mode [1/2]: 2

  Start date (YYYYMMDD): 20260330
  End date   (YYYYMMDD): 20260503
```

The same chain of scripts runs, but using the dates you entered. There's no need to edit any script or rename any file — the same date is used for the SAP extracts, the access-db output, and the SNP exceptions step.

If you mistype a date, the script will say so and ask again.

---

## What you get

After a successful run, the following files appear in `CASRA_KPI_OUTPUT/`:

| File | Description |
|------|-------------|
| `CASRA_KPI_OUTPUT_<date_from>.xlsx`        | Intermediate KPI output (before SNP exceptions). |
| `CASRA_KPI_OUTPUT_<date_from>_FINAL.xlsx`  | **Final** detail output used for reporting/visuals. Includes audit sheets for the SNP exceptions that were applied. |
| `CASRA_KPI_METRICS_<date_from>.xlsx`       | **Per-run KPI metrics** summary (one row, the metrics for this run). |
| `CASRA_KPI_METRICS_MASTER.xlsx`            | **Master KPI metrics** file. **Every run appends a new row** so you keep the full history. This is the file Power BI connects to. |

`<date_from>` is the start date you ran with, e.g. `CASRA_KPI_OUTPUT_20260501_FINAL.xlsx`.

### Run Summary (parts created + error counts)

Both intermediate and final Excel files contain a **`Run Summary`** sheet at the top. The final file's summary has everything in one row:

| Date From | Date To | Parts Created (ZMMR rows) | Rows in Output | Rows with Errors (pre-SNP) | Check_SNP errors before | Check_SNP exceptions applied | Check_SNP errors after | Rows with Errors (post-SNP) |
|---|---|---|---|---|---|---|---|---|

`Parts Created (ZMMR rows)` is the count of populated rows in the **Material Number** column of the ZMMR2199M extract — that's the "**X parts were created**" number for the KPI.

The same numbers are also printed to the console at the end of the run.

### KPI metrics file (Power BI source)

The dashboard metrics live in `CASRA_KPI_METRICS_MASTER.xlsx`. **Every run appends a new row** — full history is kept, including multiple runs on the same day. Each row carries `Report Date` (the date the script ran) plus `Date From` / `Date To` (the KPI period that run was computed for). Columns:

| Report Date | Date From | Date To | Parts Created | Storage Location | QM Insp Type | Valuation Type | Batch MNGMT | Serialized Profile | Class MOA | Unit of Measure | Hazmat | MRP Area | Total % |
|---|---|---|---|---|---|---|---|---|---|---|---|---|---|

All percentage columns are stored as **decimal values** (e.g. `0.0167` for 1.67%). Power BI is expected to format them as percentages.

Notes:
- All error counts come from the SNP-corrected output (`CASRA_KPI_OUTPUT_<date_from>_FINAL.xlsx`).
- `Parts Created` is the ZMMR2199M Material Number row count.
- `Hazmat` is currently a placeholder set to `1 / Parts Created` until the business logic is defined; the column exists today so the Power BI model doesn't have to change later.
- `Check_Class_Status` is calculated upstream but intentionally excluded from this metrics summary.

---

## Folder layout (on the machine that runs this)

```
CASRA-KPI-AUTOMATION/
├── SAP_Extracts/
│   ├── ZMNM/                 ZMNM_<date_from>.xlsx
│   └── ZMMR2199M/            ZMMR2199M_<date_from>.xlsx
├── LookUp Tables/
│   ├── RuleSloc.xlsx
│   ├── QMatMissing.xlsx
│   ├── QMATRules.xlsx
│   └── Data_Quality_ZRPN_ZGSR_NonSerialized.xlsx   (download manually each month;
│                                                    update the filename in apply_snp_exceptions.py if it changes)
└── CASRA_KPI_OUTPUT/
    ├── CASRA_KPI_OUTPUT_<date_from>.xlsx
    └── CASRA_KPI_OUTPUT_<date_from>_FINAL.xlsx
```

A `config.txt` lives next to the scripts in `src/` and holds the SAP credentials:

```
username=<sap user>
password=<sap password>
asset_num=<machine asset number>
```

---

## Running individual steps (advanced)

Each step also accepts the date range directly, which is handy for re-running just one piece:

```powershell
python access-db.py --date-from 20260501 --date-to 20260531
python apply_snp_exceptions.py --date-from 20260501 --date-to 20260531
```

If you omit the arguments, the step defaults to the previous calendar month.

## On-demand metrics from a custom input file

If you've **manually modified** an SNP-exceptions output file (corrections, what-if testing, etc.) and want to regenerate just the KPI metrics from that specific file, use `kpi-metrics-manual.py`. It runs in isolation, **does not** trigger the rest of the pipeline, and **does not** update `CASRA_KPI_METRICS_MASTER.xlsx`.

Two ways to run it:

**Interactive prompt:**

```powershell
python kpi-metrics-manual.py
```

You'll be asked for the path to the input file:

```
KPI Metrics - Manual Run
----------------------------------------
Specify the SNP-exceptions Excel file to use as input.
It must contain 'Final Output' and 'Run Summary' sheets.

Input file path:
```

**Direct path (no prompt):**

```powershell
python kpi-metrics-manual.py --input "C:\path\to\CASRA_KPI_OUTPUT_20260501_FINAL.xlsx"
```

Output: a single-row file `CASRA_KPI_OUTPUT/CASRA_KPI_METRICS_MANUAL_<today>.xlsx` containing the same KPI columns as the automated metrics file.
