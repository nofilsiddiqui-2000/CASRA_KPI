from pathlib import Path
import pandas as pd

from casra_dates import parse_date_range
from casra_paths import (
    LOOKUP_DIR,
    ensure_output_dirs,
    resolve_intermediate_output,
    snp_final_output,
)

date_from, _ = parse_date_range("apply_snp_exceptions")

# Data Quality file is downloaded manually each month and dropped into LookUp Tables.
# Update this filename if the downloaded file is named differently for a given run.
DATAQUALITY_FILE = LOOKUP_DIR / "Data_Quality_ZRPN_ZGSR_NonSerialized.xlsx"

# Column names in the DQ export vary by download; these lists are tried in order.
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


CHECK_COLUMNS = [
    "Check_Missing_Model",
    "Check_UofM",
    "Check_MOA",
    "Check_Missing_MOA_Class",
    "Check_Class_Status",
    "Check_SNP",
    "Check_Batch",
    "Check_SLoc Missing",
    "Check_SLoc_MRPInd",
    "Check_QMAT Missing",
    "Check_VType Missing",
    "Check_VType Extra",
    "Check_VType Error",
    "Check_QMAT Extra",
    "Check_MRPArea",
]


def validate_file(path: Path, label: str) -> Path:
    if not path.exists():
        raise FileNotFoundError(f"{label} file not found: {path}")
    return path


def _normalize_header(value) -> str:
    text = str(value).replace("\u00a0", " ").strip()
    return " ".join(text.split())


def find_col(df: pd.DataFrame, possible_names: list[str], label: str) -> str:
    lookup = {_normalize_header(col).lower(): col for col in df.columns}
    for name in possible_names:
        key = _normalize_header(name).lower()
        if key in lookup:
            return lookup[key]
    raise KeyError(
        f"Missing {label} column. Tried: {possible_names}\n"
        f"Columns found in file: {list(df.columns)}"
    )


def read_excel(path: Path) -> pd.DataFrame:
    df = pd.read_excel(path, dtype=object)
    df.columns = [_normalize_header(c) for c in df.columns]
    return df


def read_data_quality_file(path: Path) -> pd.DataFrame:
    """Load the DQ workbook, picking the sheet/row that has Date + Part columns."""
    xl = pd.ExcelFile(path)

    for sheet_name in xl.sheet_names:
        raw = pd.read_excel(path, sheet_name=sheet_name, header=None, dtype=object)
        max_scan = min(20, len(raw))

        for header_row in range(max_scan):
            headers = [_normalize_header(v) for v in raw.iloc[header_row].tolist()]
            probe = pd.DataFrame(raw.iloc[header_row + 1 :].values, columns=headers)
            probe = probe.dropna(how="all").reset_index(drop=True)

            if probe.empty:
                continue

            try:
                find_col(probe, DQ_DATE_COLUMNS, "Date")
                find_col(probe, DQ_PART_COLUMNS, "P/N")
                print(f"  Data Quality: sheet '{sheet_name}', header row {header_row + 1}")
                return probe
            except KeyError:
                continue

    # Last attempt: first sheet, row 0 as header (original behaviour).
    df = read_excel(path)
    find_col(df, DQ_DATE_COLUMNS, "Date")
    find_col(df, DQ_PART_COLUMNS, "P/N")
    return df


def read_run_summary(path: Path) -> pd.DataFrame:
    # access-db.py writes a "Run Summary" sheet alongside the data. If we're
    # running against an older intermediate file that doesn't have it, fall
    # back to an empty frame (we'll still record the SNP-side stats).
    try:
        return pd.read_excel(path, sheet_name="Run Summary", dtype=object)
    except (ValueError, KeyError):
        return pd.DataFrame()


def clean_part(value) -> str:
    if pd.isna(value):
        return ""

    return (
        str(value)
        .replace("\u00a0", " ")
        .strip()
        .upper()
    )


def clean_date(value) -> str:
    if pd.isna(value):
        return ""

    parsed = pd.to_datetime(value, errors="coerce")

    if pd.isna(parsed):
        return ""

    return parsed.strftime("%Y-%m-%d")


def apply_snp_exceptions(access_df: pd.DataFrame, dq_df: pd.DataFrame):
    required_access_cols = ["Material Number", "Created on", "Check_SNP", "Errors"]

    missing_access = [col for col in required_access_cols if col not in access_df.columns]
    missing_checks = [col for col in CHECK_COLUMNS if col not in access_df.columns]

    if missing_access:
        raise KeyError(f"Missing columns in Access output: {missing_access}")

    if missing_checks:
        raise KeyError(f"Missing check columns in Access output: {missing_checks}")

    dq_date_col = find_col(dq_df, DQ_DATE_COLUMNS, "Date")
    dq_part_col = find_col(dq_df, DQ_PART_COLUMNS, "P/N")

    access_df = access_df.copy()
    dq_df = dq_df.copy()

    access_df["Check_SNP"] = pd.to_numeric(access_df["Check_SNP"], errors="coerce").fillna(0).astype(int)

    access_df["_PartKey"] = access_df["Material Number"].apply(clean_part)
    access_df["_DateKey"] = access_df["Created on"].apply(clean_date)

    dq_df["_PartKey"] = dq_df[dq_part_col].apply(clean_part)
    dq_df["_DateKey"] = dq_df[dq_date_col].apply(clean_date)

    audit_cols = [c for c in DQ_AUDIT_COLUMNS if c in dq_df.columns]
    dq_export_cols = ["_PartKey", "_DateKey", dq_date_col, dq_part_col] + audit_cols

    dq_keys_df = (
        dq_df.loc[
            (dq_df["_PartKey"] != "") & (dq_df["_DateKey"] != ""),
            dq_export_cols,
        ]
        .drop_duplicates(subset=["_PartKey", "_DateKey"])
        .copy()
    )

    exception_keys = set(zip(dq_keys_df["_PartKey"], dq_keys_df["_DateKey"]))

    access_df["_SNP_Exception_Match"] = [
        (part_key, date_key) in exception_keys
        for part_key, date_key in zip(access_df["_PartKey"], access_df["_DateKey"])
    ]

    snp_before = access_df[access_df["Check_SNP"].eq(1)].copy()

    to_clear_mask = access_df["Check_SNP"].eq(1) & access_df["_SNP_Exception_Match"]

    access_df.loc[to_clear_mask, "Check_SNP"] = 0

    for col in CHECK_COLUMNS:
        access_df[col] = pd.to_numeric(access_df[col], errors="coerce").fillna(0).astype(int)

    access_df["Errors"] = access_df[CHECK_COLUMNS].sum(axis=1)

    snp_audit = snp_before[
        [
            "Material Number",
            "Created on",
            "Check_SNP",
            "_PartKey",
            "_DateKey",
            "_SNP_Exception_Match",
        ]
    ].copy()

    matched_keys = snp_before.loc[
        snp_before["_SNP_Exception_Match"],
        ["_PartKey", "_DateKey"]
    ].drop_duplicates()

    matched_dq_rows = matched_keys.merge(
        dq_keys_df,
        how="left",
        on=["_PartKey", "_DateKey"]
    )

    remaining_snp_errors = access_df[access_df["Check_SNP"].eq(1)].copy()

    final_df = access_df.drop(
        columns=["_PartKey", "_DateKey", "_SNP_Exception_Match"],
        errors="ignore"
    )

    return final_df, snp_audit, matched_dq_rows, remaining_snp_errors, dq_keys_df


def main() -> None:
    access_output_file = validate_file(
        resolve_intermediate_output(date_from),
        "Access output (Intermediate/ or legacy CASRA_KPI_OUTPUT root)",
    )
    dataquality_file = validate_file(DATAQUALITY_FILE, "Data Quality file")

    ensure_output_dirs()
    final_output_file = snp_final_output(date_from)

    access_df = read_excel(access_output_file)
    print(f"  Reading Data Quality file: {dataquality_file.name}")
    dq_df = read_data_quality_file(dataquality_file)
    run_summary = read_run_summary(access_output_file)

    final_df, snp_audit, matched_dq_rows, remaining_snp_errors, dq_keys_df = apply_snp_exceptions(access_df, dq_df)

    before_snp_errors = int(pd.to_numeric(access_df["Check_SNP"], errors="coerce").fillna(0).astype(int).sum())
    after_snp_errors = int(final_df["Check_SNP"].sum())
    exceptions_applied = before_snp_errors - after_snp_errors
    rows_with_errors_post = int((final_df["Errors"] > 0).sum())

    snp_summary_cols = {
        "Check_SNP errors before": before_snp_errors,
        "Check_SNP exceptions applied": exceptions_applied,
        "Check_SNP errors after": after_snp_errors,
        "Rows with Errors (post-SNP)": rows_with_errors_post,
    }

    if run_summary.empty:
        run_summary = pd.DataFrame([snp_summary_cols])
    else:
        for col, val in snp_summary_cols.items():
            run_summary[col] = val

    with pd.ExcelWriter(final_output_file, engine="openpyxl") as writer:
        final_df.to_excel(writer, sheet_name="Final Output", index=False)
        run_summary.to_excel(writer, sheet_name="Run Summary", index=False)
        snp_audit.to_excel(writer, sheet_name="SNP Audit", index=False)
        matched_dq_rows.to_excel(writer, sheet_name="Matched DQ Rows", index=False)
        remaining_snp_errors.to_excel(writer, sheet_name="Remaining SNP Errors", index=False)
        dq_keys_df.to_excel(writer, sheet_name="DQ Exception Keys", index=False)

    print(f"Input file: {access_output_file}")
    print(f"Data Quality file: {dataquality_file}")
    print()
    print("Run Summary:")
    if not run_summary.empty:
        for col in run_summary.columns:
            print(f"  {col}: {run_summary.iloc[0][col]}")
    print(f"\nOutput created: {final_output_file}")


if __name__ == "__main__":
    main()
