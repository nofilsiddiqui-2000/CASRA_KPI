"""Build the CASRA KPI metrics summary used for the Power BI dashboard.

Reads SNP_Final/CASRA_KPI_OUTPUT_<date_from>_<date_to>_FINAL.xlsx, writes per-run metrics
and appends to KPI_Master/CASRA_KPI_METRICS_MASTER.xlsx.
"""

from datetime import date, datetime

import pandas as pd
from openpyxl import load_workbook

from casra_common import (
    PARTS_CREATED_COL,
    REPORT_DATE_COL,
    ensure_output_dirs,
    kpi_master_output,
    kpi_metrics_output,
    parse_date_range,
    snp_final_output,
    validate_file,
    yyyymmdd_to_date,
)

DATE_COLUMNS = (REPORT_DATE_COL, "Date From", "Date To")

# KPI bucket name -> Check_* columns summed for that bucket (logic unchanged).
METRIC_CHECK_GROUPS: dict[str, list[str]] = {
    "Storage Location": ["Check_SLoc Missing", "Check_SLoc_MRPInd"],
    "QM Insp Type": ["Check_QMAT Extra", "Check_QMAT Missing"],
    "Valuation Type": ["Check_VType Extra", "Check_VType Missing", "Check_VType Error"],
    "Batch MNGMT": ["Check_Batch"],
    "Serialized Profile": ["Check_SNP"],
    "Class MOA": ["Check_MOA", "Check_Missing_Model", "Check_Missing_MOA_Class"],
    "Unit of Measure": ["Check_UofM"],
    "Hazmat": ["Check_Hazards"],
    "MRP Area": ["Check_MRPArea"],
}

# Suffix for the raw part-count column that accompanies each metric %.
COUNT_SUFFIX = " Count"


def count_column(metric_name: str) -> str:
    return f"{metric_name}{COUNT_SUFFIX}"


# Each metric bucket contributes two columns: the % and the raw part count
# (the numerator, before dividing by Parts Created). They are interleaved so
# each count sits directly next to its percentage.
METRIC_COLUMNS = [REPORT_DATE_COL, "Date From", "Date To", "Parts Created"]
for _name in METRIC_CHECK_GROUPS:
    METRIC_COLUMNS.append(_name)
    METRIC_COLUMNS.append(count_column(_name))
METRIC_COLUMNS += ["Total %", "Total Count"]


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
            "Run access_db.py and apply_snp_exceptions.py first."
        ) from exc

    if run_summary.empty:
        raise RuntimeError(f"Run Summary sheet in {final_xlsx} is empty.")
    if PARTS_CREATED_COL not in run_summary.columns:
        raise KeyError(
            f"Column '{PARTS_CREATED_COL}' not found in Run Summary. "
            f"Found: {list(run_summary.columns)}"
        )
    return int(run_summary[PARTS_CREATED_COL].iloc[0])


def compute_metrics(
    final_df: pd.DataFrame,
    parts_created: int,
    date_from: str = "",
    date_to: str = "",
) -> dict:
    if parts_created <= 0:
        raise ValueError(
            f"Parts Created is {parts_created}; cannot compute KPI percentages."
        )

    metrics: dict = {
        REPORT_DATE_COL: date.today(),
        "Date From": yyyymmdd_to_date(date_from) or pd.NaT,
        "Date To": yyyymmdd_to_date(date_to) or pd.NaT,
        "Parts Created": parts_created,
    }

    total_count = 0
    for metric_name, check_cols in METRIC_CHECK_GROUPS.items():
        count = sum(sum_check(final_df, col) for col in check_cols)
        metrics[metric_name] = count / parts_created
        metrics[count_column(metric_name)] = count
        total_count += count

    metrics["Total %"] = sum(metrics[name] for name in METRIC_CHECK_GROUPS)
    metrics["Total Count"] = total_count

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


def append_to_master(master_path, new_row: dict) -> int:
    """Add one metrics row to the master workbook. Returns the new row count.

    Only the new row is written. Everything already in the file is left exactly
    as it is - corrected numbers, HazMat errors patched in by hand, notes, extra
    columns, extra sheets, formatting.

    This used to read the whole sheet, re-normalise every date and write the
    file back out from scratch, which threw away all of that on each run.

    The new row is laid out to match the header already in the sheet, so a
    column you added or moved stays where you put it. A metric that has no
    column in the sheet is reported rather than silently dropped.
    """
    row = normalize_date_columns(pd.DataFrame([new_row], columns=METRIC_COLUMNS)).iloc[0]

    if not master_path.exists():
        pd.DataFrame([row], columns=METRIC_COLUMNS).to_excel(
            master_path, index=False, sheet_name="Metrics"
        )
        return 1

    workbook = load_workbook(master_path)
    sheet = workbook["Metrics"] if "Metrics" in workbook.sheetnames else workbook.worksheets[0]

    headers = [cell.value for cell in sheet[1]]
    if not any(headers):
        raise RuntimeError(
            f"{master_path} has no header row on sheet '{sheet.title}'; "
            "cannot tell which column is which. Fix or remove the file."
        )

    missing = [col for col in METRIC_COLUMNS if col not in headers]
    if missing:
        print(f"  WARNING: no column in the master file for: {missing}")
        print("           those values are not being written. Add the columns, or")
        print("           accept that this run's numbers for them live only in the")
        print("           per-run metrics file.")

    target_row = sheet.max_row + 1
    for index, header in enumerate(headers, start=1):
        if header not in row.index:
            continue
        value = row[header]
        cell = sheet.cell(row=target_row, column=index)
        cell.value = None if pd.isna(value) else value
        if copy_number_format(cell, sheet, target_row, index):
            cell.number_format = sheet.cell(row=target_row - 1, column=index).number_format

    workbook.save(master_path)
    return target_row - 1


def copy_number_format(cell, sheet, target_row: int, index: int) -> bool:
    """Should the new cell inherit the number format of the cell above it?

    Inheriting is what keeps a column the operator has formatted themselves -
    a percentage column, a particular date style - looking the same after a
    run. Two cases where it must not happen:

      * The row above is the header. Its format is 'General', and forcing that
        onto Report Date / Date From / Date To writes the date as a bare serial
        number (46266 instead of 2026-09-01). Power BI then reads a number
        where it expects a date and the refresh fails.

      * The row above is a data row whose cell was never formatted as a date
        ('General'), which happens when an earlier row holds the date as text.
        Same serial-number result.

    In both cases openpyxl has already given the cell a proper date format from
    the value itself, so the right move is to leave that alone.
    """
    if target_row - 1 < 2:
        return False
    above_format = sheet.cell(row=target_row - 1, column=index).number_format
    if not above_format:
        return False
    is_date = isinstance(cell.value, (datetime, date))
    return not (is_date and above_format == "General")


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

    final_xlsx = snp_final_output(date_from, date_to)
    per_run_path = kpi_metrics_output(date_from, date_to)
    master_path = kpi_master_output()

    validate_file(final_xlsx, "FINAL output")

    final_df = pd.read_excel(final_xlsx, sheet_name="Final Output")
    parts_created = get_parts_created(final_xlsx)
    metrics = compute_metrics(final_df, parts_created, date_from, date_to)

    normalize_date_columns(pd.DataFrame([metrics], columns=METRIC_COLUMNS)).to_excel(
        per_run_path, index=False, sheet_name="Metrics"
    )
    master_rows = append_to_master(master_path, metrics)

    print("\nKPI Metrics:")
    print_metrics(metrics)

    print(f"\nPer-run metrics file: {per_run_path}")
    print(f"Master metrics file:  {master_path}  ({master_rows} row(s))")


if __name__ == "__main__":
    main()