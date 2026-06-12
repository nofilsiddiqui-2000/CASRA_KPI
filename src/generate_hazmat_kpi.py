"""Apply HazMat KPI logic after SNP exceptions.

Reads SNP Final + ZMNM, appends HAZ parts from ZMNM with Check_Hazards = 1.

Final Output is built from ZMNM rows that have SAPInt (often populated via ZMMR q31).
HAZ parts typically exist only on the ZMNM extract and are not present in Final Output,
so each HAZ ZMNM row is appended with Material Number, Created on, Created, etc.

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
from casra_excel import (
    coerce_check_columns,
    find_col,
    normalize_material_key,
    read_excel_access,
    read_run_summary,
    validate_file,
)
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
HAZMAT_INDICATOR_COLUMNS = [
    "HazMat indicator",
    "HazMat Indicator",
    "Hazmat indicator",
    "Haz Mat indicator",
]
MATERIAL_NUMBER_COLUMNS = ["Material Number"]
LEGACY_ERROR_TYPE_COL = "Error Type"

_SRC = Path(__file__).parent


def _load_access_db():
    spec = importlib.util.spec_from_file_location("access_db", _SRC / "access-db.py")
    if spec is None or spec.loader is None:
        raise ImportError("Cannot load access-db.py")
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


def filter_hazmat_parts(df: pd.DataFrame, hazmat_col: str) -> pd.DataFrame:
    hazmat = df[hazmat_col].fillna("").astype("string").str.strip().str.upper()
    material_col = find_col(df, MATERIAL_NUMBER_COLUMNS, "Material Number")
    return df.loc[hazmat.eq(HAZMAT_VALUE) & df[material_col].notna()].copy()


def zmnm_field_value(zmnm_row: pd.Series, source_names: list[str]):
    for name in source_names:
        if name in zmnm_row.index and pd.notna(zmnm_row[name]):
            text = str(zmnm_row[name]).strip()
            if text:
                return zmnm_row[name]
    return pd.NA


def build_haz_row(
    zmnm_row: pd.Series,
    output_columns: list[str],
    report_date,
    query31_columns: dict[str, list[str]],
) -> dict:
    row: dict = {}
    for col in output_columns:
        if col == REPORT_DATE_COL:
            row[col] = report_date
        elif col == CHECK_HAZARDS_COL:
            row[col] = 1
        elif col in CHECK_COLUMNS:
            row[col] = 0
        elif col == "Errors":
            row[col] = 1
        elif col in query31_columns:
            row[col] = zmnm_field_value(zmnm_row, query31_columns[col])
        elif col in zmnm_row.index and pd.notna(zmnm_row[col]):
            row[col] = zmnm_row[col]
        else:
            row[col] = pd.NA
    row[CHECK_HAZARDS_COL] = 1
    row["Errors"] = 1
    return row


def materials_with_hazards_flag(final_df: pd.DataFrame, material_col: str) -> set[str]:
    keys = final_df[material_col].map(normalize_material_key)
    haz_mask = pd.to_numeric(final_df[CHECK_HAZARDS_COL], errors="coerce").fillna(0).astype(int).eq(1)
    return {key for key, flagged in zip(keys, haz_mask) if key and flagged}


def apply_hazmat_to_final(
    final_df: pd.DataFrame,
    hazmat_df: pd.DataFrame,
    query31_columns: dict[str, list[str]],
) -> tuple[pd.DataFrame, int, int]:
    """Append ZMNM HAZ rows not already represented with Check_Hazards = 1."""
    final_df = final_df.copy()
    final_df = final_df.drop(columns=[LEGACY_ERROR_TYPE_COL], errors="ignore")
    final_df = coerce_check_columns(final_df, CHECK_COLUMNS)

    material_col = find_col(final_df, MATERIAL_NUMBER_COLUMNS, "Material Number")
    haz_material_col = find_col(hazmat_df, MATERIAL_NUMBER_COLUMNS, "Material Number")

    already_haz = materials_with_hazards_flag(final_df, material_col)
    report_date = final_df[REPORT_DATE_COL].iloc[0] if REPORT_DATE_COL in final_df.columns else date.today()
    output_columns = list(final_df.columns)

    new_rows = []
    for _, zmnm_row in hazmat_df.iterrows():
        key = normalize_material_key(zmnm_row[haz_material_col])
        if not key or key in already_haz:
            continue
        new_rows.append(build_haz_row(zmnm_row, output_columns, report_date, query31_columns))
        already_haz.add(key)

    appended_rows = len(new_rows)
    if new_rows:
        final_df = pd.concat([final_df, pd.DataFrame(new_rows, columns=output_columns)], ignore_index=True)

    final_df["Errors"] = final_df[CHECK_COLUMNS].sum(axis=1)
    flagged = int(final_df[CHECK_HAZARDS_COL].sum())
    return final_df, flagged, appended_rows


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
    appended_rows: int,
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
        "HAZ rows appended": appended_rows,
        "Hazmat %": haz_pct,
    }])


def apply_hazmat_pass(final_path: Path, zmnm_path: Path) -> dict:
    access_db = _load_access_db()
    zmnm_df = read_excel_access(zmnm_path)
    zmnm_df = access_db.ensure_description_column(zmnm_df)

    hazmat_col = find_col(zmnm_df, HAZMAT_INDICATOR_COLUMNS, "HazMat indicator")
    hazmat_df = filter_hazmat_parts(zmnm_df, hazmat_col)
    haz_rows_in_zmnm = len(hazmat_df)

    final_df, run_summary, other_sheets = read_workbook_sheets(final_path)
    if PARTS_CREATED_COL not in run_summary.columns:
        raise KeyError(f"Run Summary missing '{PARTS_CREATED_COL}'. Run access-db.py first.")
    parts_created = int(run_summary[PARTS_CREATED_COL].iloc[0])

    updated_final, check_hazards_flagged, appended_rows = apply_hazmat_to_final(
        final_df, hazmat_df, access_db.QUERY31_COLUMNS
    )

    run_summary["Check_Hazards errors"] = check_hazards_flagged
    run_summary["HAZ rows appended"] = appended_rows
    run_summary["Rows with Errors (post-HAZ)"] = int((updated_final["Errors"] > 0).sum())

    write_workbook(final_path, updated_final, run_summary, other_sheets)

    debug_file = hazmat_kpi_output(date.today().strftime("%Y%m%d"))
    debug_summary = build_debug_summary(
        zmnm_path, final_path, haz_rows_in_zmnm, check_hazards_flagged, appended_rows, parts_created
    )
    with pd.ExcelWriter(debug_file, engine="openpyxl") as writer:
        hazmat_df.to_excel(writer, sheet_name="HAZ Parts", index=False)
        debug_summary.to_excel(writer, sheet_name="Summary", index=False)

    haz_pct = check_hazards_flagged / parts_created if parts_created else 0.0
    return {
        "haz_rows_in_zmnm": haz_rows_in_zmnm,
        "check_hazards_flagged": check_hazards_flagged,
        "appended_rows": appended_rows,
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
    print(f"HAZ rows appended:      {result['appended_rows']}")
    print(f"Hazmat %:               {result['haz_pct']:.4f}  ({result['haz_pct']:.2%})")
    print(f"\nDebug file:             {result['debug_file']}")
    print(f"Power BI copy:          {sharepoint_copy}")


if __name__ == "__main__":
    main()
