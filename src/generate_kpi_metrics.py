"""Build the CASRA KPI metrics summary used for the Power BI dashboard.

Reads SNP_Final/CASRA_KPI_OUTPUT_<date_from>_<date_to>_FINAL.xlsx, writes per-run metrics
and appends to KPI_Master/CASRA_KPI_METRICS_MASTER.xlsx.
"""

from datetime import date

import pandas as pd

from casra_constants import HAZ_PARTS_COL, PARTS_CREATED_COL, REPORT_DATE_COL
from casra_dates import parse_date_range, yyyymmdd_to_date
from casra_excel import validate_file
from casra_paths import (
    SHAREPOINT_KPI_MASTER_DIR,
    ensure_output_dirs,
    kpi_master_output,
    kpi_metrics_output,
    mirror_to_sharepoint,
    resolve_snp_final_output,
)

DATE_COLUMNS = (REPORT_DATE_COL, "Date From", "Date To")

METRIC_COLUMNS = [
    REPORT_DATE_COL,
    "Date From",
    "Date To",
    "Parts Created",
    "Storage Location",
    "QM Insp Type",
    "Valuation Type",
    "Batch MNGMT",
    "Serialized Profile",
    "Class MOA",
    "Unit of Measure",
    "Hazmat",
    "MRP Area",
    "Total %",
]

# KPI bucket name -> Check_* columns summed for that bucket (logic unchanged).
METRIC_CHECK_GROUPS: dict[str, list[str]] = {
    "Storage Location": ["Check_SLoc Missing", "Check_SLoc_MRPInd"],
    "QM Insp Type": ["Check_QMAT Extra", "Check_QMAT Missing"],
    "Valuation Type": ["Check_VType Extra", "Check_VType Missing", "Check_VType Error"],
    "Batch MNGMT": ["Check_Batch"],
    "Serialized Profile": ["Check_SNP"],
    "Class MOA": ["Check_MOA", "Check_Missing_Model", "Check_Missing_MOA_Class"],
    "Unit of Measure": ["Check_UofM"],
    "MRP Area": ["Check_MRPArea"],
}


def sum_check(df: pd.DataFrame, col: str) -> int:
    if col not in df.columns:
        raise KeyError(
            f"Column '{col}' not found in Final Output. "
            f"Available columns: {list(df.columns)}"
        )
    return int(pd.to_numeric(df[col], errors="coerce").fillna(0).astype(int).sum())


def get_parts_created(final_xlsx) -> int:
    try:
        run_summary = pd.read_excel(final_xlsx, sheet_name="Run Summary")
    except (ValueError, KeyError) as exc:
        raise RuntimeError(
            f"Run Summary sheet not found in {final_xlsx}. "
            "Run access-db.py and apply_snp_exceptions.py first."
        ) from exc

    if run_summary.empty:
        raise RuntimeError(f"Run Summary sheet in {final_xlsx} is empty.")
    if PARTS_CREATED_COL not in run_summary.columns:
        raise KeyError(
            f"Column '{PARTS_CREATED_COL}' not found in Run Summary. "
            f"Found: {list(run_summary.columns)}"
        )
    return int(run_summary[PARTS_CREATED_COL].iloc[0])


def get_haz_parts(final_xlsx) -> int:
    run_summary = pd.read_excel(final_xlsx, sheet_name="Run Summary", dtype=object)
    if run_summary.empty or HAZ_PARTS_COL not in run_summary.columns:
        raise RuntimeError(
            f"Run Summary in {final_xlsx} is missing '{HAZ_PARTS_COL}'. "
            "Run generate_hazmat_kpi.py first."
        )
    return int(run_summary[HAZ_PARTS_COL].iloc[0])


def compute_metrics(
    final_df: pd.DataFrame,
    parts_created: int,
    haz_parts: int,
    date_from: str = "",
    date_to: str = "",
) -> dict:
    if parts_created <= 0:
        raise ValueError(
            f"Parts Created is {parts_created}; cannot compute KPI percentages."
        )

    def pct(col_name: str) -> float:
        return sum_check(final_df, col_name) / parts_created

    metrics: dict = {
        REPORT_DATE_COL: date.today(),
        "Date From": yyyymmdd_to_date(date_from) or pd.NaT,
        "Date To": yyyymmdd_to_date(date_to) or pd.NaT,
        "Parts Created": parts_created,
    }

    for metric_name, check_cols in METRIC_CHECK_GROUPS.items():
        metrics[metric_name] = sum(pct(col) for col in check_cols)

    metrics["Hazmat"] = haz_parts / parts_created
    metrics["Total %"] = sum(metrics[name] for name in METRIC_CHECK_GROUPS) + metrics["Hazmat"]

    return metrics


def normalize_date_columns(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    for col in DATE_COLUMNS:
        if col not in df.columns:
            continue
        df[col] = df[col].apply(
            lambda v: pd.NaT if pd.isna(v) else (yyyymmdd_to_date(v) or pd.NaT)
        )
    return df


def append_to_master(master_path, new_row: dict) -> pd.DataFrame:
    new_row_df = pd.DataFrame([new_row], columns=METRIC_COLUMNS)

    if master_path.exists():
        existing = normalize_date_columns(pd.read_excel(master_path))
        combined = pd.concat([existing, new_row_df], ignore_index=True)
    else:
        combined = new_row_df

    combined = normalize_date_columns(combined.reindex(columns=METRIC_COLUMNS))
    combined.to_excel(master_path, index=False, sheet_name="Metrics")
    return combined


def print_metrics(metrics: dict) -> None:
    for col in METRIC_COLUMNS:
        value = metrics[col]
        if isinstance(value, float):
            print(f"  {col:<20} {value:.4f}  ({value:.2%})")
        else:
            print(f"  {col:<20} {value}")


def main() -> None:
    date_from, date_to = parse_date_range("generate_kpi_metrics")
    ensure_output_dirs()

    final_xlsx = resolve_snp_final_output(date_from, date_to)
    per_run_path = kpi_metrics_output(date_from, date_to)
    master_path = kpi_master_output()

    validate_file(final_xlsx, "FINAL output")

    final_df = pd.read_excel(final_xlsx, sheet_name="Final Output")
    parts_created = get_parts_created(final_xlsx)
    haz_parts = get_haz_parts(final_xlsx)
    metrics = compute_metrics(final_df, parts_created, haz_parts, date_from, date_to)

    normalize_date_columns(pd.DataFrame([metrics], columns=METRIC_COLUMNS)).to_excel(
        per_run_path, index=False, sheet_name="Metrics"
    )
    master_df = append_to_master(master_path, metrics)

    print("\nKPI Metrics:")
    print_metrics(metrics)
    sharepoint_copy = mirror_to_sharepoint(master_path, SHAREPOINT_KPI_MASTER_DIR)

    print(f"\nPer-run metrics file: {per_run_path}")
    print(f"Master metrics file:  {master_path}  ({len(master_df)} row(s))")
    print(f"Power BI copy:        {sharepoint_copy}")


if __name__ == "__main__":
    main()
