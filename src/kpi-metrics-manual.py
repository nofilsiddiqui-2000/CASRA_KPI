"""On-demand KPI metrics generation from a manually selected input file.

Use this when you've modified an SNP-exceptions output file (corrections,
what-if testing, etc.) and want to regenerate the KPI metrics summary
from that specific file. This script is NOT part of the automated
Run_KPI.py pipeline. It runs in isolation and does NOT update the
master metrics file (CASRA_KPI_METRICS_MASTER.xlsx).

Usage:
    python kpi-metrics-manual.py
        Prompts for the SNP-exceptions Excel file path.

    python kpi-metrics-manual.py --input "C:\\path\\to\\file.xlsx"
        Skips the prompt and uses the supplied file directly.

The input file must contain both a 'Final Output' sheet (with the
Check_* columns) and a 'Run Summary' sheet (with Parts Created (ZMMR
rows)). Any file produced by apply_snp_exceptions.py satisfies this.
"""

import argparse
from datetime import date
from pathlib import Path
import pandas as pd

# Reuse the exact calculation logic from the automated metrics script.
from generate_kpi_metrics import (
    METRIC_COLUMNS,
    OUTPUT_DIR,
    compute_metrics,
    get_parts_created,
    print_metrics,
    validate_file,
)


def _stringify_yyyymmdd(value) -> str:
    # Run Summary dates are written as YYYYMMDD strings. Excel may round-trip
    # them as ints/floats, so coerce safely.
    if pd.isna(value):
        return ""
    if isinstance(value, (int, float)):
        return str(int(value))
    if isinstance(value, pd.Timestamp):
        return value.strftime("%Y%m%d")
    return str(value).strip()


def get_dates_from_run_summary(input_xlsx: Path) -> tuple[str, str]:
    try:
        run_summary = pd.read_excel(input_xlsx, sheet_name="Run Summary")
    except (ValueError, KeyError):
        return "", ""

    if run_summary.empty:
        return "", ""

    date_from = (
        _stringify_yyyymmdd(run_summary["Date From"].iloc[0])
        if "Date From" in run_summary.columns else ""
    )
    date_to = (
        _stringify_yyyymmdd(run_summary["Date To"].iloc[0])
        if "Date To" in run_summary.columns else ""
    )
    return date_from, date_to


def parse_args() -> Path | None:
    parser = argparse.ArgumentParser(
        description="Manual KPI metrics generation from a chosen SNP-exceptions Excel file."
    )
    parser.add_argument(
        "--input",
        type=Path,
        default=None,
        help="Path to the SNP-exceptions Excel file (.xlsx). If omitted, you'll be prompted.",
    )
    return parser.parse_args().input


def prompt_for_path() -> Path:
    print("\nKPI Metrics - Manual Run")
    print("-" * 40)
    print("Specify the SNP-exceptions Excel file to use as input.")
    print("It must contain 'Final Output' and 'Run Summary' sheets.")

    while True:
        raw = input("\nInput file path: ").strip().strip('"').strip("'")
        if not raw:
            print("  Path cannot be empty. Try again.")
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
    cli_path = parse_args()
    input_file = cli_path if cli_path is not None else prompt_for_path()

    validate_file(input_file, "SNP-exceptions input")
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

    final_df = pd.read_excel(input_file, sheet_name="Final Output")
    parts_created = get_parts_created(input_file)
    date_from, date_to = get_dates_from_run_summary(input_file)

    metrics = compute_metrics(final_df, parts_created, date_from, date_to)

    today = date.today().strftime("%Y%m%d")
    output_file = OUTPUT_DIR / f"CASRA_KPI_METRICS_MANUAL_{today}.xlsx"

    pd.DataFrame([metrics], columns=METRIC_COLUMNS).to_excel(
        output_file, index=False, sheet_name="Metrics"
    )

    print("\nKPI Metrics (manual run):")
    print_metrics(metrics)

    print(f"\nInput file:           {input_file}")
    print(f"Output metrics file:  {output_file}")
    print("Note: master metrics file (CASRA_KPI_METRICS_MASTER.xlsx) was NOT updated for this manual run.")


if __name__ == "__main__":
    main()
