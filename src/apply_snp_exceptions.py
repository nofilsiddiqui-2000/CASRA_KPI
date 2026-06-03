import pandas as pd

from casra_constants import CHECK_COLUMNS, DQ_AUDIT_COLUMNS, DQ_DATE_COLUMNS, DQ_PART_COLUMNS
from casra_dates import parse_date_range
from casra_excel import (
    coerce_check_columns,
    find_col,
    read_data_quality_file,
    read_excel_table,
    read_run_summary,
    validate_file,
)
from casra_paths import LOOKUP_DIR, ensure_output_dirs, resolve_intermediate_output, snp_final_output

# Downloaded manually each month into LookUp Tables — update filename when it changes.
DATAQUALITY_FILE = LOOKUP_DIR / "Data_Quality_ZRPN_ZGSR_NonSerialized.xlsx"


def clean_part(value) -> str:
    if pd.isna(value):
        return ""
    return str(value).replace("\u00a0", " ").strip().upper()


def clean_date(value) -> str:
    if pd.isna(value):
        return ""
    parsed = pd.to_datetime(value, errors="coerce")
    if pd.isna(parsed):
        return ""
    return parsed.strftime("%Y-%m-%d")


def apply_snp_exceptions(access_df: pd.DataFrame, dq_df: pd.DataFrame):
    required_access = ["Material Number", "Created on", "Check_SNP", "Errors"]
    missing_access = [c for c in required_access if c not in access_df.columns]
    missing_checks = [c for c in CHECK_COLUMNS if c not in access_df.columns]

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
    dq_keys_df = (
        dq_df.loc[(dq_df["_PartKey"] != "") & (dq_df["_DateKey"] != "")]
        .drop_duplicates(subset=["_PartKey", "_DateKey"])
        [["_PartKey", "_DateKey", dq_date_col, dq_part_col] + audit_cols]
        .copy()
    )

    access_df = access_df.merge(
        dq_keys_df[["_PartKey", "_DateKey"]].assign(_SNP_Exception_Match=True),
        on=["_PartKey", "_DateKey"],
        how="left",
    )
    access_df["_SNP_Exception_Match"] = access_df["_SNP_Exception_Match"].fillna(False).astype(bool)

    snp_before = access_df[access_df["Check_SNP"].eq(1)].copy()

    access_df.loc[
        access_df["Check_SNP"].eq(1) & access_df["_SNP_Exception_Match"],
        "Check_SNP",
    ] = 0

    access_df = coerce_check_columns(access_df, CHECK_COLUMNS)
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

    matched_dq_rows = (
        snp_audit.loc[snp_audit["_SNP_Exception_Match"], ["_PartKey", "_DateKey"]]
        .drop_duplicates()
        .merge(dq_keys_df, on=["_PartKey", "_DateKey"], how="left")
    )

    remaining_snp_errors = access_df[access_df["Check_SNP"].eq(1)].copy()
    final_df = access_df.drop(columns=["_PartKey", "_DateKey", "_SNP_Exception_Match"], errors="ignore")

    return final_df, snp_audit, matched_dq_rows, remaining_snp_errors, dq_keys_df


def update_run_summary(run_summary: pd.DataFrame, snp_stats: dict) -> pd.DataFrame:
    if run_summary.empty:
        return pd.DataFrame([snp_stats])
    for col, val in snp_stats.items():
        run_summary[col] = val
    return run_summary


def main() -> None:
    date_from, _ = parse_date_range("apply_snp_exceptions")
    ensure_output_dirs()

    access_file = validate_file(
        resolve_intermediate_output(date_from),
        "Access output (Intermediate/ or legacy CASRA_KPI_OUTPUT root)",
    )
    dq_file = validate_file(DATAQUALITY_FILE, "Data Quality file")
    output_file = snp_final_output(date_from)

    access_df = read_excel_table(access_file)
    print(f"  Reading Data Quality file: {dq_file.name}")
    dq_df = read_data_quality_file(dq_file)
    run_summary = read_run_summary(access_file)

    final_df, snp_audit, matched_dq_rows, remaining_snp_errors, dq_keys_df = apply_snp_exceptions(
        access_df, dq_df
    )

    before_snp = int(access_df["Check_SNP"].sum())
    after_snp = int(final_df["Check_SNP"].sum())
    run_summary = update_run_summary(run_summary, {
        "Check_SNP errors before": before_snp,
        "Check_SNP exceptions applied": before_snp - after_snp,
        "Check_SNP errors after": after_snp,
        "Rows with Errors (post-SNP)": int((final_df["Errors"] > 0).sum()),
    })

    with pd.ExcelWriter(output_file, engine="openpyxl") as writer:
        final_df.to_excel(writer, sheet_name="Final Output", index=False)
        run_summary.to_excel(writer, sheet_name="Run Summary", index=False)
        snp_audit.to_excel(writer, sheet_name="SNP Audit", index=False)
        matched_dq_rows.to_excel(writer, sheet_name="Matched DQ Rows", index=False)
        remaining_snp_errors.to_excel(writer, sheet_name="Remaining SNP Errors", index=False)
        dq_keys_df.to_excel(writer, sheet_name="DQ Exception Keys", index=False)

    print(f"Input file: {access_file}")
    print(f"Data Quality file: {dq_file}")
    print("\nRun Summary:")
    if not run_summary.empty:
        for col in run_summary.columns:
            print(f"  {col}: {run_summary.iloc[0][col]}")
    print(f"\nOutput created: {output_file}")


if __name__ == "__main__":
    main()
