"""Build the CASRA KPI metrics summary used for the Power BI dashboard.

Inputs (all read from the FINAL Excel produced by apply_snp_exceptions.py):
    CASRA_KPI_OUTPUT/CASRA_KPI_OUTPUT_<date_from>_FINAL.xlsx
        - Final Output sheet: per-row check columns (summed for error counts).
        - Run Summary sheet:  Parts Created (ZMMR rows) used as the divisor.

Outputs:
    CASRA_KPI_OUTPUT/CASRA_KPI_METRICS_<date_from>.xlsx   (single row, per-run)
    CASRA_KPI_OUTPUT/CASRA_KPI_METRICS_MASTER.xlsx        (accumulating master;
                                                           same Report Date overwrites)

All percentage columns are stored as decimal values. Power BI is expected
to format them as percentages.

Hazmat is a placeholder until the business logic is defined; it is computed
as `1 / Parts Created` so it behaves like a single-error percentage and
keeps the Power BI schema stable.

Check_Class_Status is intentionally excluded from this metrics file (it is
still calculated upstream in access-db.py / apply_snp_exceptions.py).
"""

from datetime import date
from pathlib import Path
import pandas as pd

from casra_dates import parse_date_range


ROOT_DIR = Path(r"C:\Users\B1020000\Documents\Nofil\Dashboards\CASRA MM Dashboard\CASRA-KPI-AUTOMATION")
OUTPUT_DIR = ROOT_DIR / "CASRA_KPI_OUTPUT"

PARTS_CREATED_COL = "Parts Created (ZMMR rows)"

# Output column order, matching the spec.
METRIC_COLUMNS = [
    "Report Date",
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


def validate_file(path: Path, label: str) -> Path:
    if not path.exists():
        raise FileNotFoundError(f"{label} file not found: {path}")
    return path


def sum_check(df: pd.DataFrame, col: str) -> int:
    if col not in df.columns:
        raise KeyError(
            f"Column '{col}' not found in Final Output. "
            f"Available columns: {list(df.columns)}"
        )
    return int(pd.to_numeric(df[col], errors="coerce").fillna(0).astype(int).sum())


def get_parts_created(final_xlsx: Path) -> int:
    try:
        run_summary = pd.read_excel(final_xlsx, sheet_name="Run Summary")
    except (ValueError, KeyError) as exc:
        raise RuntimeError(
            f"Run Summary sheet not found in {final_xlsx}. Make sure access-db.py "
            "and apply_snp_exceptions.py have been run before this step."
        ) from exc

    if run_summary.empty:
        raise RuntimeError(f"Run Summary sheet in {final_xlsx} is empty.")

    if PARTS_CREATED_COL not in run_summary.columns:
        raise KeyError(
            f"Column '{PARTS_CREATED_COL}' not found in Run Summary. "
            f"Found: {list(run_summary.columns)}"
        )

    return int(run_summary[PARTS_CREATED_COL].iloc[0])


def compute_metrics(final_df: pd.DataFrame, parts_created: int) -> dict:
    if parts_created <= 0:
        raise ValueError(
            f"Parts Created is {parts_created}; cannot compute KPI percentages."
        )

    def pct(col_name: str) -> float:
        return sum_check(final_df, col_name) / parts_created

    storage_location = pct("Check_SLoc Missing") + pct("Check_SLoc_MRPInd")
    qm_insp_type = pct("Check_QMAT Extra") + pct("Check_QMAT Missing")
    valuation_type = (
        pct("Check_VType Extra")
        + pct("Check_VType Missing")
        + pct("Check_VType Error")
    )
    batch_mngmt = pct("Check_Batch")
    serialized_profile = pct("Check_SNP")
    class_moa = (
        pct("Check_MOA")
        + pct("Check_Missing_Model")
        + pct("Check_Missing_MOA_Class")
    )
    unit_of_measure = pct("Check_UofM")
    hazmat = 1 / parts_created  # placeholder until business logic is defined.
    mrp_area = pct("Check_MRPArea")

    total_pct = (
        storage_location
        + qm_insp_type
        + valuation_type
        + batch_mngmt
        + serialized_profile
        + class_moa
        + unit_of_measure
        + hazmat
        + mrp_area
    )

    return {
        "Report Date": date.today(),
        "Parts Created": parts_created,
        "Storage Location": storage_location,
        "QM Insp Type": qm_insp_type,
        "Valuation Type": valuation_type,
        "Batch MNGMT": batch_mngmt,
        "Serialized Profile": serialized_profile,
        "Class MOA": class_moa,
        "Unit of Measure": unit_of_measure,
        "Hazmat": hazmat,
        "MRP Area": mrp_area,
        "Total %": total_pct,
    }


def upsert_master(master_path: Path, new_row: dict) -> pd.DataFrame:
    new_row_df = pd.DataFrame([new_row], columns=METRIC_COLUMNS)

    if master_path.exists():
        existing = pd.read_excel(master_path)
        if "Report Date" in existing.columns:
            existing["Report Date"] = pd.to_datetime(
                existing["Report Date"], errors="coerce"
            ).dt.date
            existing = existing[existing["Report Date"] != new_row["Report Date"]]
        combined = pd.concat([existing, new_row_df], ignore_index=True)
    else:
        combined = new_row_df

    combined = combined.sort_values("Report Date").reset_index(drop=True)
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
    date_from, _ = parse_date_range("generate_kpi_metrics")

    final_xlsx = OUTPUT_DIR / f"CASRA_KPI_OUTPUT_{date_from}_FINAL.xlsx"
    per_run_metrics = OUTPUT_DIR / f"CASRA_KPI_METRICS_{date_from}.xlsx"
    master_metrics = OUTPUT_DIR / "CASRA_KPI_METRICS_MASTER.xlsx"

    validate_file(final_xlsx, "FINAL output")
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

    final_df = pd.read_excel(final_xlsx, sheet_name="Final Output")
    parts_created = get_parts_created(final_xlsx)

    metrics = compute_metrics(final_df, parts_created)

    per_run_df = pd.DataFrame([metrics], columns=METRIC_COLUMNS)
    per_run_df.to_excel(per_run_metrics, index=False, sheet_name="Metrics")

    master_df = upsert_master(master_metrics, metrics)

    print("\nKPI Metrics:")
    print_metrics(metrics)

    print(f"\nPer-run metrics file: {per_run_metrics}")
    print(f"Master metrics file:  {master_metrics}  ({len(master_df)} row(s))")


if __name__ == "__main__":
    main()
