# CASRA KPI — Business Guide

This document explains **what** the CASRA KPI process does and **how** each error flag and dashboard metric is derived. It is written for business analysts and process owners.

For how to run the automation (scripts, folders, Power BI paths), see [README.md](README.md).

---

## What this process delivers

Each run produces:

1. A **part-level detail file** — one row per material created in the period, with yes/no flags for each type of data-quality issue.
2. A **monthly KPI summary** — high-level percentages for leadership dashboards (Power BI).

The goal is to measure how many newly created parts have master-data issues, by category (storage location, SNP, valuation type, etc.), and to track that over time.

---

## End-to-end flow (business view)

```text
1. SAP extraction (date range you choose)
      ZMNM  →  list of materials created in the period
      ZMMR2199M  →  detailed change records (which fields were touched per part)

2. KPI build (access-db logic)
      Combines ZMNM + ZMMR + lookup tables
      →  flags each part with Check_* columns (0 = OK, 1 = issue found)
      →  Errors = total number of issue types flagged for that part

3. SNP exceptions
      Known-good parts from the monthly Data Quality file can clear a false SNP flag

4. KPI metrics
      Rolls up the final file into dashboard percentages (Parts Created, Storage Location %, etc.)
```

---

## Step 1 — SAP data extraction

You choose the reporting period when you start the run:

- **Automated mode** — previous calendar month (e.g. 1 May–31 May if you run in June).
- **Manual mode** — any start and end date you enter.

### ZMNM transaction

- Pulls materials **created** in that date range (with your standard user filters in SAP).
- This is the **base list of parts** the KPI is about — “how many parts were created this period.”

### ZMMR2199M transaction

- Pulls **change/detail records** for the same date range.
- Each row describes something that was checked or changed for a part (which SAP field, which report type, actual value, plant, material type, etc.).
- Only **HALB** material type rows are kept in the extract used for KPIs.
- The same part can appear on **many rows** in ZMMR (one row per field/report combination). Those rows are used only to **detect issues** on each part — not to count how many parts were created.

---

## Step 2 — Building the part-level output (access-db logic)

### Starting point

- The process starts from the **ZMNM** extract (materials created in the period).
- It then **adds** material information from ZMMR where those parts also appear there (material number, description, created date, plant, group, etc.).
- The result is one working table: **one row per part** from ZMNM (enriched from ZMMR).

### How ZMMR is used to find issues

ZMMR is filtered to rows that have a valid **Counter** and **Material Number** (real data rows, not blanks).

For each check below, the logic asks: *“Does this part appear on a ZMMR row that matches this condition?”*

- If **yes** → that check column is set to **1** (issue flagged).
- If **no** → **0** (no issue for that check on this part).

All checks are independent. A part can have several checks set to 1 at the same time.

---

## Check columns — what each one means

| Column | Business meaning | When is it flagged (1)? |
|--------|------------------|-------------------------|
| **Check_Missing_Model** | Model / MOI-related master data issue | ZMMR row where the SAP **Field** is `zmm_moi_pn-admoi` |
| **Check_UofM** | Unit of measure issue | ZMMR row where **Field** is `mara-meins` |
| **Check_MOA** | MOA (version) issue | ZMMR row where **Field** is `klah-versi` |
| **Check_Missing_MOA_Class** | MOA class missing | ZMMR row where **Field Name** is `moa class` |
| **Check_Class_Status** | Class status issue | ZMMR row where **Field** is `klah-statu ` (matches legacy Access spelling, including trailing space) |
| **Check_SNP** | Serialized profile (SNP) issue | ZMMR row where **Field** is `marc-sernp` |
| **Check_Batch** | Batch management issue | ZMMR row where **Field** is one of: `mara-mhdrz`, `mara-mhdhb`, `marc-xchpf` |
| **Check_SLoc Missing** | Storage location missing (rule-based) | ZMMR **Plant SLoc** report, **Field Name** = `storage location`, and the combination of material type + item category + plant + actual value matches a **Key** in the **RuleSloc** lookup table |
| **Check_SLoc_MRPInd** | Storage location MRP indicator issue | ZMMR **Plant SLoc** report and **Field** = `mard-diskz` |
| **Check_QMAT Missing** | QM inspection type missing | Part number appears in the **QMatMissing** lookup list |
| **Check_VType Missing** | Valuation type missing (wrong scenario) | ZMMR matches one of: new valuation type + `mvke-mtpos`; or HALB + “valuation type no value”; or HALB (not NORM/ERLA) + core/rotable valuation type + `mvke-mtpos` |
| **Check_VType Extra** | Valuation type present when it should not be | ZMMR matches one of: non-HALB + “valuation type no value”; or HALB with NORM and ERLA together on core/rotable (legacy rule); or any “valuation type” report with **Field** = `mara-mtart` |
| **Check_VType Error** | Valuation type data error | Any “valuation type” report with **Field** = `mbew-bklas` or `mbew-vprsv` **and** an actual value filled in |
| **Check_QMAT Extra** | QM inspection type not in allowed rules | ZMMR **Plant data** report, **Field** = `qmat-art`, has an actual value, but the built rule **Key** is **not** in the **QMATRules** lookup |
| **Check_MRPArea** | MRP area issue | ZMMR row where **Field** is `marc-diber` |

### Errors column

**Errors** = sum of all Check_* columns above (each contributes 0 or 1).

- **0** → no issue types flagged for that part.
- **1 or more** → at least one category failed; higher numbers mean more categories flagged on the same part.

### What appears in the final detail file

Only parts with **SAPInt** populated are exported (parts tied to the integration / reporting scope). The file includes material attributes plus all Check_* columns and **Errors**.

### Parts Created (Run Summary)

Counted from **ZMNM**: number of rows with a **Material Number** filled in (simple row count from the ZMNM extract). This is the denominator for dashboard percentages (“X parts were created this period”).

---

## Lookup tables (supporting rules)

| Lookup file | Used for |
|-------------|----------|
| **RuleSloc.xlsx** | Keys that define which storage-location situations count as “SLoc missing” |
| **QMatMissing.xlsx** | List of parts that should be flagged as missing QM inspection type |
| **QMATRules.xlsx** | Allowed QM inspection type keys — parts outside this set can be flagged as “QMAT Extra” |

These sit in the **LookUp Tables** folder and are maintained by the business (not re-downloaded from SAP each month).

---

## Step 3 — SNP exceptions

After the KPI build, some parts may still show **Check_SNP = 1**.

Each month you download **Data_Quality_ZRPN_ZGSR_NonSerialized** (or equivalent) into **LookUp Tables**. That file lists parts that are **known exceptions** — they may look wrong in ZMMR but are accepted by the business.

For each part in the KPI output:

- If **Check_SNP = 1** **and** the same **Material Number** and **Created on** date appear in the Data Quality file  
  → **Check_SNP** is set to **0** (exception applied).  
- **Errors** is recalculated after that change.

Parts that still have **Check_SNP = 1** after this step are true SNP errors for reporting.

The **FINAL** detail file (used for Power BI part-level analysis) includes audit sheets showing what was cleared and what remained.

---

## Step 4 — Dashboard KPI metrics

The metrics file does **not** list every part. It produces **one summary row per run** with percentages for leadership.

| Metric | How it is calculated |
|--------|----------------------|
| **Report Date** | Date the automation was run |
| **Date From / Date To** | The SAP period you selected |
| **Parts Created** | ZMNM row count with Material Number filled in (same as Run Summary) |
| **Storage Location** | % of parts flagged for SLoc Missing + % flagged for SLoc MRP indicator |
| **QM Insp Type** | % QMAT Extra + % QMAT Missing |
| **Valuation Type** | % VType Extra + % VType Missing + % VType Error |
| **Batch MNGMT** | % Check_Batch |
| **Serialized Profile** | % Check_SNP (after SNP exceptions) |
| **Class MOA** | % Check_MOA + % Missing Model + % Missing MOA Class |
| **Unit of Measure** | % Check_UofM |
| **Hazmat** | Placeholder until business rules are defined (currently a fixed small value so the dashboard column exists) |
| **MRP Area** | % Check_MRPArea |
| **Total %** | Sum of the metrics above |

Each percentage = *(number of parts with that check = 1) ÷ Parts Created*.

**Note:** **Check_Class_Status** is calculated in the detail file but is **not** included in these dashboard buckets (by design).

The **KPI Master** file keeps **one row per run** so Power BI can show trends over time.

---

## Power BI — which files to use

| Purpose | File |
|---------|------|
| Trend / summary dashboard | **KPI_Master** — `CASRA_KPI_METRICS_MASTER.xlsx` |
| Part-level detail / drill-down | **SNP_Final** — `CASRA_KPI_OUTPUT_<period start>_FINAL.xlsx` (one file per period) |

Both are copied to your SharePoint-synced folder after each run (see README for path configuration).

---

## Quick reference — 0 vs 1

| Value | Meaning |
|-------|---------|
| **0** on a Check_* column | No issue detected for that category on this part |
| **1** on a Check_* column | Issue detected for that category on this part |
| **Errors** | Count of how many Check_* columns are 1 for that part |

---

## Document map

| Document | Audience | Content |
|----------|----------|---------|
| **README.md** | People running the tool | Scripts, folders, how to run, Power BI paths |
| **CASRA KPI.md** (this file) | Analysts & process owners | Business rules, flags, metrics definitions |
