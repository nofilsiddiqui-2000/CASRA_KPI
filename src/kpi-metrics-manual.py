"""Manual KPI test run from user-supplied SAP extracts and a Data Quality file.

Fill in the paths below, then run:
    python kpi-metrics-manual.py

Runs access-db checks and SNP exceptions on your files, then writes metrics.
Does not update KPI_Master/CASRA_KPI_METRICS_MASTER.xlsx.
"""

import importlib.util
from datetime import date
from pathlib import Path

import pandas as pd

from apply_snp_exceptions import run_snp_pass
from casra_common import ensure_output_dirs, kpi_metrics_manual_output
from generate_kpi_metrics import (
    METRIC_COLUMNS,
    compute_metrics,
    normalize_date_columns,
    print_metrics,
)

# --- Manual test inputs (fill these in before running) ---
ZMNM_FILE = r""
ZMMR_FILE = r""
DATA_QUALITY_FILE = r""

_SRC = Path(__file__).parent


def _load_access_db():
    spec = importlib.util.spec_from_file_location("access_db", _SRC / "access-db.py")
    if spec is None or spec.loader is None:
        raise ImportError("Cannot load access-db.py")
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


def require_path(value: str, name: str) -> Path:
    text = value.strip().strip('"').strip("'")
    if not text:
        raise ValueError(f"Set {name} at the top of kpi-metrics-manual.py before running.")
    path = Path(text)
    if not path.exists():
        raise FileNotFoundError(f"{name} file not found: {path}")
    if path.suffix.lower() != ".xlsx":
        raise ValueError(f"{name} must point to an .xlsx file: {path}")
    return path


def write_detail_output(
    output_path: Path,
    final_df: pd.DataFrame,
    run_summary: pd.DataFrame,
    snp_audit: pd.DataFrame,
    matched_dq_rows: pd.DataFrame,
    remaining_snp_errors: pd.DataFrame,
    dq_keys_df: pd.DataFrame,
) -> None:
    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        final_df.to_excel(writer, sheet_name="Final Output", index=False)
        run_summary.to_excel(writer, sheet_name="Run Summary", index=False)
        snp_audit.to_excel(writer, sheet_name="SNP Audit", index=False)
        matched_dq_rows.to_excel(writer, sheet_name="Matched DQ Rows", index=False)
        remaining_snp_errors.to_excel(writer, sheet_name="Remaining SNP Errors", index=False)
        dq_keys_df.to_excel(writer, sheet_name="DQ Exception Keys", index=False)


def main() -> None:
    zmnm_path = require_path(ZMNM_FILE, "ZMNM_FILE")
    zmmr_path = require_path(ZMMR_FILE, "ZMMR_FILE")
    dq_path = require_path(DATA_QUALITY_FILE, "DATA_QUALITY_FILE")

    access_db = _load_access_db()
    ensure_output_dirs()

    paths = access_db.build_paths_from_files(zmnm_path, zmmr_path)
    access_output, run_summary, parts_created = access_db.run_kpi_build(paths)

    print(f"ZMNM file:              {zmnm_path}")
    print(f"ZMMR file:              {zmmr_path}")
    print(f"Parts created (ZMNM):   {parts_created}")
    print(f"Rows in output:           {len(access_output)}")

    final_df, run_summary, snp_audit, matched_dq_rows, remaining_snp_errors, dq_keys_df = run_snp_pass(
        access_output, run_summary, dq_path
    )

    metrics = compute_metrics(final_df, parts_created)
    run_date = date.today().strftime("%Y%m%d")
    metrics_file = kpi_metrics_manual_output(run_date)
    detail_file = metrics_file.with_name(f"CASRA_KPI_MANUAL_{run_date}_FINAL.xlsx")

    normalize_date_columns(pd.DataFrame([metrics], columns=METRIC_COLUMNS)).to_excel(
        metrics_file, index=False, sheet_name="Metrics"
    )
    write_detail_output(
        detail_file,
        final_df,
        run_summary,
        snp_audit,
        matched_dq_rows,
        remaining_snp_errors,
        dq_keys_df,
    )

    print("\nKPI Metrics (manual run):")
    print_metrics(metrics)
    print(f"\nDetail output file:     {detail_file}")
    print(f"Metrics output file:    {metrics_file}")
    print("Note: KPI_Master/CASRA_KPI_METRICS_MASTER.xlsx was NOT updated.")


if __name__ == "__main__":
    main()
