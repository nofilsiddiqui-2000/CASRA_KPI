"""Apply HazMat KPI logic after SNP exceptions.

Reads SNP Final + ZMNM, flags HAZ parts with Check_Hazards = 1 (same as other check
columns), backfills Created on / Created from ZMNM, recalculates Errors, and writes a
debug workbook.

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
    CHECK_HAZARDS_COL,
    PARTS_CREATED_COL,
    REPORT_DATE_COL,
)
from casra_dates import parse_date_range
from casra_excel import coerce_check_columns, find_col, read_excel_access, read_run_summary, validate_file
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
CREATED_ON_COLUMNS = ["Created on", "Created On"]
CREATED_COLUMNS = ["Created", "Created By"]
LEGACY_ERROR_TYPE_COL = "Error Type"

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


def build_haz_lookup(hazmat_df: pd.DataFrame) -> tuple[set[str], dict[str, dict[str, object]]]:
    material_col = find_col(hazmat_df, MATERIAL_NUMBER_COLUMNS, "Material Number")
    created_on_col = find_col(hazmat_df, CREATED_ON_COLUMNS, "Created on")
    created_col = find_col(hazmat_df, CREATED_COLUMNS, "Created")

    haz_materials: set[str] = set()
    lookup: dict[str, dict[str, object]] = {}
    for _, row in hazmat_df.iterrows():
        material = norm_material(row[material_col])
        if not material:
            continue
        haz_materials.add(material)
        if material not in lookup:
            lookup[material] = {
                "Created on": row[created_on_col],
                "Created": row[created_col],
            }
    return haz_materials, lookup


def apply_hazards_check(
    final_df: pd.DataFrame,
    haz_materials: set[str],
    zmnm_lookup: dict[str, dict[str, object]],
) -> tuple[pd.DataFrame, int]:
    final_df = final_df.copy()
    final_df = final_df.drop(columns=[LEGACY_ERROR_TYPE_COL], errors="ignore")
    final_df = coerce_check_columns(final_df, CHECK_COLUMNS)

    material_col = find_col(final_df, MATERIAL_NUMBER_COLUMNS, "Material Number")
    materials = final_df[material_col].map(norm_material)
    haz_mask = materials.isin(haz_materials)
    final_df.loc[haz_mask, CHECK_HAZARDS_COL] = 1

    for idx in final_df.index[haz_mask]:
        material = norm_material(final_df.at[idx, material_col])
        source = zmnm_lookup.get(material, {})
        for col in ("Created on", "Created"):
            value = source.get(col)
            if value is not None and not pd.isna(value) and str(value).strip():
                final_df.at[idx, col] = value

    final_df["Errors"] = final_df[CHECK_COLUMNS].sum(axis=1)
    flagged = int(final_df[CHECK_HAZARDS_COL].sum())
    return final_df, flagged


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
    haz_rows_in_zmnm: int,
    check_hazards_flagged: int,
    parts_created: int,
) -> pd.DataFrame:
    haz_pct = check_hazards_flagged / parts_created if parts_created else 0.0
    return pd.DataFrame([{
        REPORT_DATE_COL: date.today(),
        "Source ZMNM File": zmnm_path.name,
        "Source Final File": final_path.name,
        PARTS_CREATED_COL: parts_created,
        "HAZ rows in ZMNM": haz_rows_in_zmnm,
        "Check_Hazards flagged": check_hazards_flagged,
        "Hazmat %": haz_pct,
    }])


def apply_hazmat_pass(final_path: Path, zmnm_path: Path) -> dict:
    access_db = _load_access_db()
    zmnm_df = read_excel_access(zmnm_path)
    zmnm_df = access_db.ensure_description_column(zmnm_df)

    hazmat_col = find_col(zmnm_df, HAZMAT_INDICATOR_COLUMNS, "HazMat indicator")
    hazmat_df = filter_hazmat_parts(zmnm_df, hazmat_col)
    haz_rows_in_zmnm = len(hazmat_df)
    haz_materials, zmnm_lookup = build_haz_lookup(hazmat_df)

    final_df, run_summary, other_sheets = read_workbook_sheets(final_path)
    if PARTS_CREATED_COL not in run_summary.columns:
        raise KeyError(f"Run Summary missing '{PARTS_CREATED_COL}'. Run access-db.py first.")
    parts_created = int(run_summary[PARTS_CREATED_COL].iloc[0])

    updated_final, check_hazards_flagged = apply_hazards_check(final_df, haz_materials, zmnm_lookup)

    run_summary["Check_Hazards errors"] = check_hazards_flagged
    run_summary["Rows with Errors (post-HAZ)"] = int((updated_final["Errors"] > 0).sum())

    write_workbook(final_path, updated_final, run_summary, other_sheets)

    debug_file = hazmat_kpi_output(date.today().strftime("%Y%m%d"))
    debug_summary = build_debug_summary(
        zmnm_path, final_path, haz_rows_in_zmnm, check_hazards_flagged, parts_created
    )
    with pd.ExcelWriter(debug_file, engine="openpyxl") as writer:
        hazmat_df.to_excel(writer, sheet_name="HAZ Parts", index=False)
        debug_summary.to_excel(writer, sheet_name="Summary", index=False)

    haz_pct = check_hazards_flagged / parts_created if parts_created else 0.0
    return {
        "haz_rows_in_zmnm": haz_rows_in_zmnm,
        "check_hazards_flagged": check_hazards_flagged,
        "parts_created": parts_created,
        "haz_pct": haz_pct,
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
    print(f"HAZ rows in ZMNM:       {result['haz_rows_in_zmnm']}")
    print(f"Check_Hazards flagged:  {result['check_hazards_flagged']}")
    print(f"Hazmat %:               {result['haz_pct']:.4f}  ({result['haz_pct']:.2%})")
    print(f"\nDebug file:             {result['debug_file']}")
    print(f"Power BI copy:          {sharepoint_copy}")


if __name__ == "__main__":
    main()
