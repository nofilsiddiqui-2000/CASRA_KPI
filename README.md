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

After a successful run, two files appear in `CASRA_KPI_OUTPUT/`:

| File | Description |
|------|-------------|
| `CASRA_KPI_OUTPUT_<date_from>.xlsx`        | Intermediate KPI output (before SNP exceptions). |
| `CASRA_KPI_OUTPUT_<date_from>_FINAL.xlsx`  | **Final** output used for reporting/visuals. Includes audit sheets for the SNP exceptions that were applied. |

`<date_from>` is the start date you ran with, e.g. `CASRA_KPI_OUTPUT_20260501_FINAL.xlsx`.

### Run Summary (parts created + error counts)

Both Excel files contain a **`Run Summary`** sheet at the top. The final file's summary has everything in one row:

| Date From | Date To | Parts Created (ZMMR rows) | Rows in Output | Rows with Errors (pre-SNP) | Check_SNP errors before | Check_SNP exceptions applied | Check_SNP errors after | Rows with Errors (post-SNP) |
|---|---|---|---|---|---|---|---|---|

`Parts Created (ZMMR rows)` is the count of populated rows in the **Material Number** column of the ZMMR2199M extract — that's the "**X parts were created**" number for the KPI.

The same numbers are also printed to the console at the end of the run.

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
