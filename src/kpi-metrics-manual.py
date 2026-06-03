"""On-demand KPI metrics from a manually chosen SNP-exceptions Excel file.

Does not update KPI_Master/CASRA_KPI_METRICS_MASTER.xlsx.

Usage:
    python kpi-metrics-manual.py
    python kpi-metrics-manual.py --input "path\\to\\file.xlsx"
"""

import argparse
from datetime import date
from pathlib import Path

import pandas as pd

from casra_excel import period_from_run_summary, validate_file
from casra_paths import ensure_output_dirs, kpi_metrics_manual_output
from generate_kpi_metrics import (
    METRIC_COLUMNS,
    compute_metrics,
    get_parts_created,
    normalize_date_columns,
    print_metrics,
)


def parse_args() -> Path | None:
    parser = argparse.ArgumentParser(
        description="Manual KPI metrics from a chosen SNP-exceptions Excel file."
    )
    parser.add_argument("--input", type=Path, default=None, help="Path to .xlsx input file.")
    return parser.parse_args().input


def prompt_for_path() -> Path:
    print("\nKPI Metrics - Manual Run")
    print("-" * 40)
    print("Specify the SNP-exceptions Excel file (Final Output + Run Summary sheets).")

    while True:
        raw = input("\nInput file path: ").strip().strip('"').strip("'")
        if not raw:
            print("  Path cannot be empty.")
            continue
        path = Path(raw)
        if not path.exists():
            print(f"  File not found: {path}")
            continue
        if path.suffix.lower() != ".xlsx":
            print("  File must be an .xlsx workbook.")
            continue
        return path


def main() -> None:
    input_file = parse_args() or prompt_for_path()
    validate_file(input_file, "SNP-exceptions input")
    ensure_output_dirs()

    final_df = pd.read_excel(input_file, sheet_name="Final Output")
    parts_created = get_parts_created(input_file)
    date_from, date_to = period_from_run_summary(input_file)

    metrics = compute_metrics(final_df, parts_created, date_from, date_to)
    output_file = kpi_metrics_manual_output(date.today().strftime("%Y%m%d"))

    normalize_date_columns(pd.DataFrame([metrics], columns=METRIC_COLUMNS)).to_excel(
        output_file, index=False, sheet_name="Metrics"
    )

    print("\nKPI Metrics (manual run):")
    print_metrics(metrics)
    print(f"\nInput file:           {input_file}")
    print(f"Output metrics file:  {output_file}")
    print("Note: KPI_Master/CASRA_KPI_METRICS_MASTER.xlsx was NOT updated.")


if __name__ == "__main__":
    main()
