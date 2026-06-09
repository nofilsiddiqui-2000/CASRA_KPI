"""Shared column lists used across the KPI pipeline."""

# Columns that contribute to the per-row Errors total (order matches Access q47).
CHECK_COLUMNS = [
    "Check_SNP",
    "Check_UofM",
    "Check_QMAT Extra",
    "Check_QMAT Missing",
    "Check_VType Extra",
    "Check_VType Missing",
    "Check_VType Error",
    "Check_SLoc Missing",
    "Check_SLoc_MRPInd",
    "Check_Batch",
    "Check_MOA",
    "Check_Missing_Model",
    "Check_Missing_MOA_Class",
    "Check_Class_Status",
    "Check_MRPArea",
]

PARTS_CREATED_COL = "Parts Created (ZMNM rows)"
REPORT_DATE_COL = "Report Date"

# Data Quality workbook — column names vary by monthly download.
DQ_DATE_COLUMNS = [
    "Date",
    "DATE",
    "Created on",
    "Created On",
    "Created Date",
    "Creation Date",
]
DQ_PART_COLUMNS = [
    "P/N",
    "P / N",
    "PN",
    "P-N",
    "Part Number",
    "Part No",
    "Part No.",
    "Material Number",
    "Material",
]
DQ_AUDIT_COLUMNS = [
    "Item Category Group",
    "REF",
    "Comment",
    "By",
]
