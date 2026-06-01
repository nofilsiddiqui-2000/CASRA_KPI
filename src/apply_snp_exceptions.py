from pathlib import Path
import pandas as pd


ROOT_DIR = Path(r"C:\Users\B1020000\Documents\Nofil\Dashboards\CASRA MM Dashboard\CASRA-KPI-AUTOMATION")
OUTPUT_DIR = ROOT_DIR / "CASRA_KPI_OUTPUT"

ACCESS_OUTPUT_FILE = Path(r"C:\Users\B1020000\Documents\Nofil\Dashboards\CASRA MM Dashboard\CASRA-KPI-AUTOMATION\Check_SNP\CASRA_KPI_OUTPUT_MANUAL_20260529.xlsx")
DATAQUALITY_FILE = Path(r"C:\Users\B1020000\Documents\Nofil\Dashboards\CASRA MM Dashboard\CASRA-KPI-AUTOMATION\Check_SNP\Data_Quality_ZRPN_ZGSR_NonSerialized.xlsx")

FINAL_OUTPUT_FILE = OUTPUT_DIR / "CASRA_KPI_OUTPUT_SNP_EXCEPTIONS.xlsx"


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


def read_excel(path: Path) -> pd.DataFrame:
    df = pd.read_excel(path, dtype=object)
    df.columns = df.columns.astype(str).str.strip()
    return df


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
    required_dq_cols = ["Date", "P/N"]

    missing_access = [col for col in required_access_cols if col not in access_df.columns]
    missing_dq = [col for col in required_dq_cols if col not in dq_df.columns]
    missing_checks = [col for col in CHECK_COLUMNS if col not in access_df.columns]

    if missing_access:
        raise KeyError(f"Missing columns in Access output: {missing_access}")

    if missing_dq:
        raise KeyError(f"Missing columns in Data Quality file: {missing_dq}")

    if missing_checks:
        raise KeyError(f"Missing check columns in Access output: {missing_checks}")

    access_df = access_df.copy()
    dq_df = dq_df.copy()

    access_df["Check_SNP"] = pd.to_numeric(access_df["Check_SNP"], errors="coerce").fillna(0).astype(int)

    access_df["_PartKey"] = access_df["Material Number"].apply(clean_part)
    access_df["_DateKey"] = access_df["Created on"].apply(clean_date)

    dq_df["_PartKey"] = dq_df["P/N"].apply(clean_part)
    dq_df["_DateKey"] = dq_df["Date"].apply(clean_date)

    dq_keys_df = (
        dq_df.loc[
            (dq_df["_PartKey"] != "") & (dq_df["_DateKey"] != ""),
            ["_PartKey", "_DateKey", "Date", "P/N", "Item Category Group", "REF", "Comment", "By"]
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
    access_output_file = validate_file(ACCESS_OUTPUT_FILE, "Access output")
    dataquality_file = validate_file(DATAQUALITY_FILE, "Data Quality file")

    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

    access_df = read_excel(access_output_file)
    dq_df = read_excel(dataquality_file)

    final_df, snp_audit, matched_dq_rows, remaining_snp_errors, dq_keys_df = apply_snp_exceptions(access_df, dq_df)

    before_snp_errors = int(pd.to_numeric(access_df["Check_SNP"], errors="coerce").fillna(0).astype(int).sum())
    after_snp_errors = int(final_df["Check_SNP"].sum())
    exceptions_applied = before_snp_errors - after_snp_errors

    with pd.ExcelWriter(FINAL_OUTPUT_FILE, engine="openpyxl") as writer:
        final_df.to_excel(writer, sheet_name="Final Output", index=False)
        snp_audit.to_excel(writer, sheet_name="SNP Audit", index=False)
        matched_dq_rows.to_excel(writer, sheet_name="Matched DQ Rows", index=False)
        remaining_snp_errors.to_excel(writer, sheet_name="Remaining SNP Errors", index=False)
        dq_keys_df.to_excel(writer, sheet_name="DQ Exception Keys", index=False)

    print(f"Input file: {access_output_file}")
    print(f"Data Quality file: {dataquality_file}")
    print(f"Check_SNP errors before exceptions: {before_snp_errors}")
    print(f"Check_SNP exceptions applied: {exceptions_applied}")
    print(f"Check_SNP errors after exceptions: {after_snp_errors}")
    print(f"Output created: {FINAL_OUTPUT_FILE}")


if __name__ == "__main__":
    main()
