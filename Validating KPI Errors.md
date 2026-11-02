# Validating KPI Errors

How to prove, part by part, why a KPI error was counted. No theory - each entry
says which file to open, what to search for, and what makes it an error.

Every condition below is **case-insensitive** but **not** whitespace-trimmed.
See [Two traps](#two-traps) at the bottom.

---

## Start here

**1. Open the final detail file:**

```
CASRA_KPI_OUTPUT/SNP_Final/CASRA_KPI_OUTPUT_<date_from>_<date_to>_FINAL.xlsx
```

Sheet **Final Output**. One row per part, with a `Check_*` column per error
type and an `Errors` column holding their sum.

**2. Find the part.** Filter `Material Number`. Whichever `Check_*` column is
**1** is the error you need to explain.

**3. Go to that column's section below.**

| KPI metric on the dashboard | Check columns behind it |
|---|---|
| Storage Location | `Check_SLoc Missing`, `Check_SLoc_MRPInd` |
| QM Insp Type | `Check_QMAT Missing`, `Check_QMAT Extra` |
| Valuation Type | `Check_VType Missing`, `Check_VType Extra`, `Check_VType Error` |
| Batch MNGMT | `Check_Batch` |
| Serialized Profile | `Check_SNP` |
| Class MOA | `Check_MOA`, `Check_Missing_Model`, `Check_Missing_MOA_Class` |
| Unit of Measure | `Check_UofM` |
| Hazmat | `Check_Hazards` |
| MRP Area | `Check_MRPArea` |

So "Valuation Type = 3" means three parts have a 1 in one of those three
columns. Filter each in turn to find which.

---

## The evidence file

Unless a section says otherwise, the proof is a row in the **ZMMR2199M
extract**:

```
SAP_Extracts/ZMMR2199M/ZMMR2199M_<date_from>.xlsx
```

Filter `Material Number` to your part. A part usually has **many** rows here,
one per field/report combination. You are looking for one row that matches the
condition. If it is there, the error is correct.

Columns used below: `Field`, `Field Name`, `Report selection` (SAP exports this
as `Rep. sel.`; the extract script renames it), `Actuel Value`, `MTyp`,
`ItCGr`, `Plnt`.

---

## Storage Location

### Check_SLoc Missing

| | |
|---|---|
| **Open** | ZMMR2199M extract, then `LookUp Tables/RuleSloc.xlsx` |
| **Find** | ZMMR rows for the part where `Report selection` = `PLANT SLOC` **and** `Field Name` = `Storage Location` |
| **Then** | Join that row's four cells into one string, in this order, with nothing between them: `MTyp` + `ItCGr` + `Plnt` + `Actuel Value`. Example: `HALB` + `ERLA` + `3099` + `0001` = `HALBERLA30990001` |
| **Error when** | That string **is present** in the `Key` column of `RuleSloc.xlsx` |

Note the direction: the key being **found** is what makes it an error. This is
the opposite of `Check_QMAT Extra` below.

### Check_SLoc_MRPInd

| | |
|---|---|
| **Open** | ZMMR2199M extract |
| **Error when** | The part has a row with `Report selection` = `PLANT SLOC` **and** `Field` = `MARD-DISKZ` |

---

## QM Insp Type

### Check_QMAT Missing

| | |
|---|---|
| **Open** | `LookUp Tables/QMatMissing.xlsx` |
| **Error when** | The part appears in the `Material Number` column |

No ZMMR row is involved. This list is the whole test - if the part is on it,
the flag is set.

### Check_QMAT Extra

| | |
|---|---|
| **Open** | ZMMR2199M extract, then `LookUp Tables/QMATRules.xlsx` |
| **Find** | ZMMR rows for the part where `Report selection` = `PLANT DATA` **and** `Field` = `QMAT-ART` **and** `Actuel Value` is not blank |
| **Then** | Build the key: `MTyp` + (`ERLA` if `ItCGr` is `ERLA`, otherwise the literal text `BLANK`) + `Plnt` + `Actuel Value`. Example: MTyp `HALB`, ItCGr `NORM`, Plnt `3099`, Actuel Value `0140` = `HALBBLANK30990140` |
| **Error when** | That key is **not** in the `Key` column of `QMATRules.xlsx` |

`ItCGr` only ever becomes `ERLA` or `BLANK` in the key. Anything that is not
`ERLA` - `NORM`, blank, anything else - becomes `BLANK`.

---

## Valuation Type

All three look at rows whose `Report selection` **starts with** `VALUATION
TYPE` (so `VALUATION TYPE NEW`, `... CORE`, `... ROTABLE`, `... NO VALUE`).

### Check_VType Missing

**Open** the ZMMR2199M extract. It is an error when the part has a row with
`Field` = `MVKE-MTPOS` **and** any one of:

- `Report selection` = `VALUATION TYPE NEW` (any MTyp, any ItCGr), or
- `Report selection` = `VALUATION TYPE NO VALUE` **and** `MTyp` = `HALB`, or
- `Report selection` = `VALUATION TYPE CORE` **and** `MTyp` = `HALB` **and** `ItCGr` is neither `NORM` nor `ERLA`, or
- `Report selection` = `VALUATION TYPE ROTABLE` **and** `MTyp` = `HALB` **and** `ItCGr` is neither `NORM` nor `ERLA`

### Check_VType Extra

**Open** the ZMMR2199M extract. It is an error when the part has a row
matching either of:

- `Report selection` starts with `VALUATION TYPE` **and** `Field` = `MARA-MTART`, or
- `Report selection` = `VALUATION TYPE NO VALUE` **and** `Field` = `MVKE-MTPOS` **and** `MTyp` is **not** `HALB`

In practice it is nearly always the first one. The code carries two more
conditions from the original Access query that require `ItCGr` to be both
`NORM` and `ERLA` at once, so they can never match. They are kept so the logic
stays identical to Access - ignore them.

### Check_VType Error

| | |
|---|---|
| **Open** | ZMMR2199M extract |
| **Error when** | The part has a row where `Report selection` starts with `VALUATION TYPE` **and** `Field` is `MBEW-BKLAS` or `MBEW-VPRSV` **and** `Actuel Value` is **not blank** |

A blank `Actuel Value` on such a row is not an error.

---

## Batch MNGMT

### Check_Batch

| | |
|---|---|
| **Open** | ZMMR2199M extract |
| **Error when** | The part has a row where `Field` is any of `MARA-MHDRZ`, `MARA-MHDHB`, `MARC-XCHPF` |

Any one of the three is enough, and matching several still counts as one error
for the part.

---

## Serialized Profile

### Check_SNP

Two stages. The flag is raised, then possibly cleared.

**Raised:**

| | |
|---|---|
| **Open** | ZMMR2199M extract |
| **Flagged when** | The part has a row with `Field` = `MARC-SERNP` |

**Cleared** if the part has an accepted exception in the Data Quality
workbook:

| | |
|---|---|
| **Open** | `LookUp Tables/Data Quality-<month>.xlsx` |
| **Cleared when** | A row exists whose `P/N` matches the part **and** whose `Date` matches the part's `Created on` from Final Output |

Both must match. Part numbers are compared trimmed and upper-cased; dates are
compared as calendar dates.

**Fastest way to check:** the FINAL workbook already shows this. Two sheets:

- **SNP Audit** - every part that was flagged, with `_SNP_Exception_Match`
  `TRUE` or `FALSE`. `FALSE` means it stayed an error.
- **Matched DQ Rows** - for the cleared ones, the actual Data Quality row that
  excused it, including `REF`, `Comment` and `By`. That is your answer when
  someone asks who approved it.

---

## Class MOA

Three separate columns, all simple `Field` / `Field Name` matches in the
ZMMR2199M extract.

| Column | Error when the part has a row with |
|---|---|
| `Check_MOA` | `Field` = `KLAH-VERSI` |
| `Check_Missing_Model` | `Field` = `ZMM_MOI_PN-ADMOI` |
| `Check_Missing_MOA_Class` | `Field Name` = `MOA Class` |

Note `Check_Missing_MOA_Class` is the only one of the three that reads **Field
Name**, not **Field**.

### Check_Class_Status

Not part of the Class MOA metric - it feeds `Errors` only.

| | |
|---|---|
| **Open** | ZMMR2199M extract |
| **Error when** | The part has a row where `Field` = `KLAH-STATU ` - **with one trailing space** |

A row with `KLAH-STATU` and no trailing space does **not** flag. See
[Two traps](#two-traps).

---

## Unit of Measure

### Check_UofM

| | |
|---|---|
| **Open** | ZMMR2199M extract |
| **Error when** | The part has a row with `Field` = `MARA-MEINS` |

---

## MRP Area

### Check_MRPArea

| | |
|---|---|
| **Open** | ZMMR2199M extract |
| **Error when** | The part has a row with `Field` = `MARC-DIBER` |

---

## Hazmat

### Check_Hazards

**While `HAZMAT_MODE = "skip"` in `generate_hazmat_kpi.py`, this is always 0
and Hazmat % is always 0.00%.** Nothing is checked. Any real HazMat error has
to be found by hand and patched into the master metrics file.

To validate by hand:

| | |
|---|---|
| **1. Get the part list** | `SAP_Extracts/ZMNM/ZMNM_<date_from>.xlsx`, filter `HazMat indicator` = `HAZ`. Or open the debug workbook `CASRA_KPI_OUTPUT/HazMat_KPI/CASRA_HAZMAT_KPI_<run date>.xlsx`, sheet **HAZ Parts** - same list |
| **2. In SAP** | Transaction **MM03**, enter the part, open the **Sales Text** tab |
| **3. Org levels** | Plant `3099`, distribution channel `00`, and check **all three** sales orgs: `3000`, `4200`, `1000` |
| **Not an error** | Every one of the three sales orgs reads `HAZARDOUS MATERIAL` |
| **Error** | Any one of them reads anything else - blank, `HAZMAT`, extra wording, a typo |

One error per part, however many sales orgs are wrong. A part you cannot open
at all is neither - record it and follow up.

When the live mode is working again, the debug workbook's **MM03 Results**
sheet will hold the text read per part per sales org, and **Skipped Parts**
the ones it could not read.

---

## Two traps

**1. Trailing spaces are not stripped.** Values are lower-cased before
comparison but never trimmed. A `Field` of `MARA-MEINS` with a trailing space
will **not** match and will not be flagged. If a part looks like it should
have been flagged and was not, check the cell for stray spaces - in Excel,
compare `LEN()` of the cell against the expected length.

The one deliberate exception is `Check_Class_Status`, which matches
`KLAH-STATU ` **with** a trailing space, because that is what the original
Access query did.

**2. The two key lookups are case-sensitive.** The strings built for
`Check_SLoc Missing` and `Check_QMAT Extra` are compared to the lookup tables
exactly as they are, with no case conversion. SAP exports `MTyp`, `ItCGr` and
`Plnt` upper-cased and the lookup tables store them upper-cased, so they line
up - but a hand-edited lookup row in mixed case will silently stop matching.

---

## Counts and totals

| | |
|---|---|
| `Errors` (per row, Final Output) | Sum of all 16 `Check_*` columns for that part |
| `Parts Created` | Number of rows in the ZMNM extract with a `Material Number` filled in. This is the denominator for every percentage |
| A metric's `%` | That metric's `Count` divided by `Parts Created` |
| `Total %` | The nine metric percentages added together |

A part can appear on more than one row of Final Output - once from the normal
build and once appended by the HazMat step - so the `Count` columns count
flagged rows, not distinct part numbers.

---

## Files, all relative to `CASRA-KPI-AUTOMATION/`

| Purpose | Path |
|---|---|
| The flags per part | `CASRA_KPI_OUTPUT/SNP_Final/CASRA_KPI_OUTPUT_<from>_<to>_FINAL.xlsx` |
| Evidence for almost every check | `SAP_Extracts/ZMMR2199M/ZMMR2199M_<from>.xlsx` |
| HAZ part list, Parts Created | `SAP_Extracts/ZMNM/ZMNM_<from>.xlsx` |
| SNP exceptions | `LookUp Tables/Data Quality-<month>.xlsx` |
| Storage location rules | `LookUp Tables/RuleSloc.xlsx` |
| QM inspection type rules | `LookUp Tables/QMATRules.xlsx` |
| QM inspection type missing list | `LookUp Tables/QMatMissing.xlsx` |
| HazMat debug sheets | `CASRA_KPI_OUTPUT/HazMat_KPI/CASRA_HAZMAT_KPI_<run date>.xlsx` |
| The metric percentages | `CASRA_KPI_OUTPUT/KPI_Master/CASRA_KPI_METRICS_MASTER.xlsx` |
