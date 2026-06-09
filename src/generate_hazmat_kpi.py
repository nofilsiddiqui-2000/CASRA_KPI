"""Standalone HazMat KPI from a ZMNM extract.

Fill in the path below, then run:
    python generate_hazmat_kpi.py

Filters ZMNM rows where HazMat indicator = HAZ and writes a summary + detail file.
Not part of Run_KPI.py yet.
"""

from datetime import date
from pathlib import Path

import pandas as pd

from casra_excel import find_col, read_excel_access, validate_file
from casra_paths import ensure_output_dirs, hazmat_kpi_output

# --- Input (fill this in before running) ---
ZMNM_FILE = r""

HAZMAT_VALUE = "HAZ"
HAZMAT_INDICATOR_COLUMNS = ["HazMat indicator", "Hazmat indicator", "Haz Mat indicator"]
MATERIAL_NUMBER_COLUMNS = ["Material Number"]


def require_path(value: str, name: str) -> Path:
    text = value.strip().strip('"').strip("'")
    if not text:
        raise ValueError(f"Set {name} at the top of generate_hazmat_kpi.py before running.")
    path = Path(text)
    validate_file(path, name)
    if path.suffix.lower() != ".xlsx":
        raise ValueError(f"{name} must point to an .xlsx file: {path}")
    return path


def count_material_rows(df: pd.DataFrame, material_col: str) -> int:
    return int(df[material_col].notna().sum())


def filter_hazmat_parts(df: pd.DataFrame, hazmat_col: str) -> pd.DataFrame:
    hazmat = df[hazmat_col].fillna("").astype("string").str.strip().str.upper()
    return df.loc[hazmat.eq(HAZMAT_VALUE)].copy()


def build_summary(
    zmnm_rows: int,
    hazmat_rows: int,
    source_file: Path,
) -> pd.DataFrame:
    hazmat_pct = hazmat_rows / zmnm_rows if zmnm_rows else 0.0
    return pd.DataFrame([{
        "Report Date": date.today(),
        "Source File": source_file.name,
        "ZMNM Rows (Material Number)": zmnm_rows,
        "HAZ Parts": hazmat_rows,
        "Hazmat %": hazmat_pct,
    }])


def main() -> None:
    zmnm_path = require_path(ZMNM_FILE, "ZMNM_FILE")
    ensure_output_dirs()

    zmnm_df = read_excel_access(zmnm_path)
    material_col = find_col(zmnm_df, MATERIAL_NUMBER_COLUMNS, "Material Number")
    hazmat_col = find_col(zmnm_df, HAZMAT_INDICATOR_COLUMNS, "HazMat indicator")

    zmnm_rows = count_material_rows(zmnm_df, material_col)
    hazmat_df = filter_hazmat_parts(zmnm_df, hazmat_col)
    hazmat_rows = len(hazmat_df)

    summary = build_summary(zmnm_rows, hazmat_rows, zmnm_path)
    output_file = hazmat_kpi_output(date.today().strftime("%Y%m%d"))

    with pd.ExcelWriter(output_file, engine="openpyxl") as writer:
        hazmat_df.to_excel(writer, sheet_name="HAZ Parts", index=False)
        summary.to_excel(writer, sheet_name="Summary", index=False)

    print(f"ZMNM file:              {zmnm_path}")
    print(f"HazMat indicator col:   {hazmat_col}")
    print(f"ZMNM rows (Material #): {zmnm_rows}")
    print(f"HAZ parts:              {hazmat_rows}")
    print(f"Hazmat %:               {summary['Hazmat %'].iloc[0]:.4f}  ({summary['Hazmat %'].iloc[0]:.2%})")
    print(f"\nOutput file:            {output_file}")


if __name__ == "__main__":
    main()
