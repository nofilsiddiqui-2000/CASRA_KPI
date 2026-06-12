"""Apply HazMat KPI logic after SNP exceptions.

Reads SNP Final + ZMNM, counts HAZ parts (HazMat indicator = HAZ), appends missing
HAZ rows to Final Output, updates Run Summary, and writes a debug workbook.

Pipeline order:
    apply_snp_exceptions  ->  generate_hazmat_kpi  ->  generate_kpi_metrics
"""

from __future__ import annotations

import importlib.util
from datetime import date
from pathlib import Path

import pandas as pd

from casra_constants import (
    CHECK_COLUMNS,
    ERROR_TYPE_COL,
    HAZMAT_ERROR_TYPE,
    HAZ_PARTS_COL,
    PARTS_CREATED_COL,
    REPORT_DATE_COL,
)
from casra_dates import parse_date_range
from casra_excel import find_col, read_excel_access, read_run_summary, validate_file
from casra_paths import (
    SAP_DIR,
    SHAREPOINT_SNP_FINAL_DIR,
    ensure_output_dirs,
    hazmat_kpi_output,
    mirror_to_sharepoint,
    resolve_snp_final_output,
)

# --- Optional manual override (leave blank for pipeline run) ---
ZMNM_FILE = r""

HAZMAT_VALUE = "HAZ"
HAZMAT_INDICATOR_COLUMNS = ["HazMat indicator", "Hazmat indicator", "Haz Mat indicator"]
MATERIAL_NUMBER_COLUMNS = ["Material Number"]

_SRC = Path(__file__).parent


def _load_access_db():
    spec = importlib.util.spec_from_file_location("access_db", _SRC / "access-db.py")
    if spec is None or spec.loader is None:
        raise ImportError("Cannot load access-db.py")
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


def norm_material(value) -> str:
    if pd.isna(value):
        return ""
    return str(value).strip()


def filter_hazmat_parts(df: pd.DataFrame, hazmat_col: str) -> pd.DataFrame:
    hazmat = df[hazmat_col].fillna("").astype("string").str.strip().str.upper()
    material_col = find_col(df, MATERIAL_NUMBER_COLUMNS, "Material Number")
    return df.loc[hazmat.eq(HAZMAT_VALUE) & df[material_col].notna()].copy()


def hazmat_materials(final_df: pd.DataFrame, material_col: str) -> set[str]:
    if ERROR_TYPE_COL not in final_df.columns:
        return set()
    mask = final_df[ERROR_TYPE_COL].fillna("").astype("string").str.strip().str.upper().eq(HAZMAT_ERROR_TYPE)
    return {norm_material(v) for v in final_df.loc[mask, material_col] if norm_material(v)}


def map_zmnm_row(
    zmnm_row: pd.Series,
    output_columns: list[str],
    report_date,
) -> dict:
    row: dict = {}
    for col in output_columns:
        if col == ERROR_TYPE_COL:
            row[col] = HAZMAT_ERROR_TYPE
        elif col in CHECK_COLUMNS or col == "Errors":
            row[col] = 0
        elif col == REPORT_DATE_COL:
            row[col] = report_date
        elif col in zmnm_row.index and pd.notna(zmnm_row[col]):
            row[col] = zmnm_row[col]
        else:
            row[col] = pd.NA
    return row


def append_hazmat_rows(final_df: pd.DataFrame, hazmat_df: pd.DataFrame) -> tuple[pd.DataFrame, int]:
    material_col = find_col(final_df, MATERIAL_NUMBER_COLUMNS, "Material Number")
    if ERROR_TYPE_COL not in final_df.columns:
        final_df = final_df.copy()
        final_df[ERROR_TYPE_COL] = pd.NA

    report_date = final_df[REPORT_DATE_COL].iloc[0] if REPORT_DATE_COL in final_df.columns else date.today()
    already_hazmat = hazmat_materials(final_df, material_col)
    output_columns = list(final_df.columns)

    new_rows = []
    for _, zmnm_row in hazmat_df.iterrows():
        material = norm_material(zmnm_row[material_col])
        if not material or material in already_hazmat:
            continue
        new_rows.append(map_zmnm_row(zmnm_row, output_columns, report_date))
        already_hazmat.add(material)

    if not new_rows:
        return final_df, 0

    appended = pd.DataFrame(new_rows, columns=output_columns)
    return pd.concat([final_df, appended], ignore_index=True), len(new_rows)


def read_workbook_sheets(path: Path) -> tuple[pd.DataFrame, pd.DataFrame, dict[str, pd.DataFrame]]:
    xl = pd.ExcelFile(path)
    final_df = pd.read_excel(path, sheet_name="Final Output", dtype=object)
    run_summary = read_run_summary(path)
    other_sheets = {
        name: pd.read_excel(path, sheet_name=name, dtype=object)
        for name in xl.sheet_names
        if name not in {"Final Output", "Run Summary"}
    }
    return final_df, run_summary, other_sheets


def write_workbook(
    path: Path,
    final_df: pd.DataFrame,
    run_summary: pd.DataFrame,
    other_sheets: dict[str, pd.DataFrame],
) -> None:
    with pd.ExcelWriter(path, engine="openpyxl") as writer:
        final_df.to_excel(writer, sheet_name="Final Output", index=False)
        run_summary.to_excel(writer, sheet_name="Run Summary", index=False)
        for sheet_name, sheet_df in other_sheets.items():
            sheet_df.to_excel(writer, sheet_name=sheet_name, index=False)


def resolve_zmnm_path(date_from: str) -> Path:
    if ZMNM_FILE.strip():
        return validate_file(Path(ZMNM_FILE.strip().strip('"').strip("'")), "ZMNM")
    return validate_file(SAP_DIR / "ZMNM" / f"ZMNM_{date_from}.xlsx", "ZMNM")


def build_debug_summary(
    zmnm_path: Path,
    final_path: Path,
    haz_parts: int,
    parts_created: int,
    appended_rows: int,
) -> pd.DataFrame:
    haz_pct = haz_parts / parts_created if parts_created else 0.0
    return pd.DataFrame([{
        REPORT_DATE_COL: date.today(),
        "Source ZMNM File": zmnm_path.name,
        "Source Final File": final_path.name,
        PARTS_CREATED_COL: parts_created,
        HAZ_PARTS_COL: haz_parts,
        "HAZ Rows Appended": appended_rows,
        "Hazmat %": haz_pct,
    }])


def apply_hazmat_pass(final_path: Path, zmnm_path: Path) -> dict:
    access_db = _load_access_db()
    zmnm_df = read_excel_access(zmnm_path)
    zmnm_df = access_db.ensure_description_column(zmnm_df)

    hazmat_col = find_col(zmnm_df, HAZMAT_INDICATOR_COLUMNS, "HazMat indicator")
    hazmat_df = filter_hazmat_parts(zmnm_df, hazmat_col)
    haz_parts = len(hazmat_df)

    final_df, run_summary, other_sheets = read_workbook_sheets(final_path)
    if PARTS_CREATED_COL not in run_summary.columns:
        raise KeyError(f"Run Summary missing '{PARTS_CREATED_COL}'. Run access-db.py first.")
    parts_created = int(run_summary[PARTS_CREATED_COL].iloc[0])

    updated_final, appended_rows = append_hazmat_rows(final_df, hazmat_df)

    run_summary[HAZ_PARTS_COL] = haz_parts
    run_summary["HAZ Rows Appended"] = appended_rows
    run_summary["Rows in Output (post-HAZ)"] = len(updated_final)

    write_workbook(final_path, updated_final, run_summary, other_sheets)

    debug_file = hazmat_kpi_output(date.today().strftime("%Y%m%d"))
    debug_summary = build_debug_summary(zmnm_path, final_path, haz_parts, parts_created, appended_rows)
    with pd.ExcelWriter(debug_file, engine="openpyxl") as writer:
        hazmat_df.to_excel(writer, sheet_name="HAZ Parts", index=False)
        debug_summary.to_excel(writer, sheet_name="Summary", index=False)

    return {
        "haz_parts": haz_parts,
        "parts_created": parts_created,
        "appended_rows": appended_rows,
        "haz_pct": haz_parts / parts_created if parts_created else 0.0,
        "final_path": final_path,
        "debug_file": debug_file,
    }


def main() -> None:
    date_from, date_to = parse_date_range("generate_hazmat_kpi")
    ensure_output_dirs()

    final_path = validate_file(resolve_snp_final_output(date_from, date_to), "SNP Final output")
    zmnm_path = resolve_zmnm_path(date_from)

    result = apply_hazmat_pass(final_path, zmnm_path)
    sharepoint_copy = mirror_to_sharepoint(final_path, SHAREPOINT_SNP_FINAL_DIR)

    print(f"ZMNM file:              {zmnm_path}")
    print(f"SNP Final file:         {result['final_path']}")
    print(f"Parts created:          {result['parts_created']}")
    print(f"HAZ parts:              {result['haz_parts']}")
    print(f"HAZ rows appended:      {result['appended_rows']}")
    print(f"Hazmat %:               {result['haz_pct']:.4f}  ({result['haz_pct']:.2%})")
    print(f"\nDebug file:             {result['debug_file']}")
    print(f"Power BI copy:          {sharepoint_copy}")


if __name__ == "__main__":
    main()
