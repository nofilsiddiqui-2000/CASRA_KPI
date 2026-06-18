"""Shared helpers for the CASRA KPI pipeline.

This module consolidates what used to live in five separate files:
    casra_constants  ->  Shared column lists
    casra_dates      ->  Date parsing / range utilities
    casra_config     ->  config.txt reader (SAP credentials)
    casra_paths      ->  Output folder layout + SharePoint sync
    casra_excel      ->  Excel read/write helpers

The pipeline scripts (access-db, apply_snp_exceptions, generate_hazmat_kpi,
generate_kpi_metrics, the SAP extracts, Run_KPI) all import from here.

Sections are ordered by dependency: constants and dates are defined first
because the Excel helpers at the bottom rely on them.
"""

from __future__ import annotations

import argparse
import os
import shutil
from datetime import date, datetime, timedelta
from pathlib import Path

import pandas as pd


# ============================================================================
# Constants  (was casra_constants.py)
# ============================================================================

# Columns that contribute to the per-row Errors total (order matches Access q47).
CHECK_COLUMNS = [
    "Check_SNP",
    "Check_UofM",
    "Check_QMAT Extra",
    "Check_QMAT Missing",
    "Check_VType Extra",
    "Check_VType Missing",
    "Check_VType Error",
    "Check_SLoc Missing",
    "Check_SLoc_MRPInd",
    "Check_Batch",
    "Check_MOA",
    "Check_Missing_Model",
    "Check_Missing_MOA_Class",
    "Check_Class_Status",
    "Check_MRPArea",
    "Check_Hazards",
]

CHECK_HAZARDS_COL = "Check_Hazards"
PARTS_CREATED_COL = "Parts Created (ZMNM rows)"
REPORT_DATE_COL = "Report Date"

# Data Quality workbook — column names vary by monthly download.
DQ_DATE_COLUMNS = [
    "Date",
    "DATE",
    "Created on",
    "Created On",
    "Created Date",
    "Creation Date",
]
DQ_PART_COLUMNS = [
    "P/N",
    "P / N",
    "PN",
    "P-N",
    "Part Number",
    "Part No",
    "Part No.",
    "Material Number",
    "Material",
]
DQ_AUDIT_COLUMNS = [
    "Item Category Group",
    "REF",
    "Comment",
    "By",
]


# ============================================================================
# Dates  (was casra_dates.py)
# ============================================================================
#
# Every step of the pipeline (SAP extracts, access-db, apply_snp_exceptions)
# agrees on a single date range. Each script accepts:
#
#     --date-from YYYYMMDD
#     --date-to   YYYYMMDD
#
# When omitted, both default to the previous calendar month. This way each
# script still runs standalone, while Run_KPI.py can drive the whole chain
# with one consistent range.

DATE_FORMAT = "%Y%m%d"


def previous_month_range() -> tuple[str, str]:
    """Return (date_from, date_to) for the previous calendar month, YYYYMMDD."""
    first_of_this_month = date.today().replace(day=1)
    last_of_prev_month = first_of_this_month - timedelta(days=1)
    first_of_prev_month = last_of_prev_month.replace(day=1)
    return (
        first_of_prev_month.strftime(DATE_FORMAT),
        last_of_prev_month.strftime(DATE_FORMAT),
    )


def _validate_yyyymmdd(value: str) -> str:
    try:
        datetime.strptime(value, DATE_FORMAT)
    except ValueError as exc:
        raise argparse.ArgumentTypeError(
            f"Invalid date '{value}'. Expected YYYYMMDD (e.g. 20260330)."
        ) from exc
    return value


def parse_date_range(description: str = "CASRA KPI step") -> tuple[str, str]:
    """Parse --date-from / --date-to from sys.argv with prev-month fallback.

    Both args are accepted by every step so Run_KPI.py can pass the same
    pair of arguments to all child scripts uniformly. Steps that only
    care about date_from can simply ignore date_to.
    """
    default_from, default_to = previous_month_range()
    parser = argparse.ArgumentParser(description=description)
    parser.add_argument(
        "--date-from",
        type=_validate_yyyymmdd,
        default=default_from,
        help=f"Start date (YYYYMMDD). Default: {default_from} (start of previous month).",
    )
    parser.add_argument(
        "--date-to",
        type=_validate_yyyymmdd,
        default=default_to,
        help=f"End date (YYYYMMDD). Default: {default_to} (end of previous month).",
    )
    args = parser.parse_args()
    return args.date_from, args.date_to


def yyyymmdd_to_date(value) -> date | None:
    """Convert YYYYMMDD (string or Excel int) to a calendar date for Excel output."""
    if value is None:
        return None

    if hasattr(value, "to_pydatetime"):
        try:
            value = value.to_pydatetime()
        except (ValueError, TypeError):
            return None

    if isinstance(value, date) and not isinstance(value, datetime):
        return value
    if isinstance(value, datetime):
        return value.date()

    text = str(value).strip()
    if not text or text.lower() in ("nan", "nat", "none"):
        return None

    # Excel sometimes stores 20260531 as float 20260531.0
    if "." in text:
        text = text.split(".", 1)[0]

    if len(text) >= 8 and text[:8].isdigit():
        try:
            return datetime.strptime(text[:8], DATE_FORMAT).date()
        except ValueError:
            return None

    try:
        parsed = datetime.fromisoformat(text[:10])
        return parsed.date()
    except ValueError:
        return None


# ============================================================================
# Config  (was casra_config.py)
# ============================================================================
#
# Reads config.txt next to the pipeline scripts (SAP credentials, etc.).


def read_config(config_dir: Path | None = None) -> dict[str, str]:
    base = config_dir or Path(__file__).resolve().parent
    config_path = base / "config.txt"
    config_data: dict[str, str] = {}

    with open(config_path, encoding="utf-8") as file:
        for line in file:
            line = line.strip()
            if not line or line.startswith("#"):
                continue
            if "==" in line:
                key, value = line.split("==", 1)
            elif "=" in line:
                key, value = line.split("=", 1)
            else:
                continue
            config_data[key.strip()] = value.strip().strip('"').strip("'")

    return config_data


# ============================================================================
# Paths  (was casra_paths.py)
# ============================================================================
#
# Folder layout for CASRA_KPI_OUTPUT. Each pipeline stage's files live in their
# own subfolder (same idea as SAP_Extracts/ZMNM and SAP_Extracts/ZMMR2199M).

ROOT_DIR = Path(
    r"C:\Users\B1020000\Documents\Nofil\Dashboards\CASRA MM Dashboard\CASRA-KPI-AUTOMATION"
)
# ROOT_DIR = Path(__file__).resolve().parent.parent

OUTPUT_ROOT = ROOT_DIR / "CASRA_KPI_OUTPUT"

# access-db.py — KPI output before SNP exceptions
INTERMEDIATE_DIR = OUTPUT_ROOT / "Intermediate"

# apply_snp_exceptions.py — final detail output after SNP exceptions
SNP_FINAL_DIR = OUTPUT_ROOT / "SNP_Final"

# generate_kpi_metrics.py — one-row metrics workbook per run
KPI_METRICS_DIR = OUTPUT_ROOT / "KPI_Metrics"

# kpi-metrics-manual.py — on-demand metrics (does not update master)
KPI_METRICS_MANUAL_DIR = OUTPUT_ROOT / "KPI_Metrics_Manual"

# generate_hazmat_kpi.py — HazMat KPI debug workbook
HAZMAT_KPI_DIR = OUTPUT_ROOT / "HazMat_KPI"

# generate_kpi_metrics.py — accumulating history for Power BI
KPI_MASTER_DIR = OUTPUT_ROOT / "KPI_Master"

LOOKUP_DIR = ROOT_DIR / "LookUp Tables"
SAP_DIR = ROOT_DIR / "SAP_Extracts"

# Power BI reads from this SharePoint-synced folder on the work desktop.
# Create SNP_Final/ and KPI_Master/ inside it, then update this path to match.
SHAREPOINT_SYNC_ROOT = Path(
    r"C:\Users\B1020000\Bombardier\SharePoint-Sync\CASRA_KPI_PowerBI"
)

SHAREPOINT_SNP_FINAL_DIR = SHAREPOINT_SYNC_ROOT / "SNP_Final"
SHAREPOINT_KPI_MASTER_DIR = SHAREPOINT_SYNC_ROOT / "KPI_Master"


def ensure_output_dirs() -> None:
    for folder in (
        INTERMEDIATE_DIR,
        SNP_FINAL_DIR,
        KPI_METRICS_DIR,
        KPI_METRICS_MANUAL_DIR,
        HAZMAT_KPI_DIR,
        KPI_MASTER_DIR,
        SHAREPOINT_SNP_FINAL_DIR,
        SHAREPOINT_KPI_MASTER_DIR,
    ):
        folder.mkdir(parents=True, exist_ok=True)


def mirror_to_sharepoint(local_file: Path, sharepoint_dir: Path) -> Path:
    """Copy a pipeline output file into the SharePoint-synced folder for Power BI."""
    if not local_file.exists():
        raise FileNotFoundError(f"Cannot mirror missing file: {local_file}")

    sharepoint_dir.mkdir(parents=True, exist_ok=True)
    destination = sharepoint_dir / local_file.name
    shutil.copy2(local_file, destination)
    return destination


def output_period_suffix(date_from: str, date_to: str) -> str:
    return f"{date_from}_{date_to}"


def intermediate_output(date_from: str, date_to: str) -> Path:
    suffix = output_period_suffix(date_from, date_to)
    return INTERMEDIATE_DIR / f"CASRA_KPI_OUTPUT_{suffix}.xlsx"


def snp_final_output(date_from: str, date_to: str) -> Path:
    suffix = output_period_suffix(date_from, date_to)
    return SNP_FINAL_DIR / f"CASRA_KPI_OUTPUT_{suffix}_FINAL.xlsx"


def resolve_intermediate_output(date_from: str, date_to: str) -> Path:
    """Prefer Intermediate/; fall back to legacy single-date filenames."""
    current = intermediate_output(date_from, date_to)
    if current.exists():
        return current
    legacy = INTERMEDIATE_DIR / f"CASRA_KPI_OUTPUT_{date_from}.xlsx"
    if legacy.exists():
        return legacy
    flat_legacy = OUTPUT_ROOT / f"CASRA_KPI_OUTPUT_{date_from}.xlsx"
    if flat_legacy.exists():
        return flat_legacy
    return current


def resolve_snp_final_output(date_from: str, date_to: str) -> Path:
    """Prefer SNP_Final/; fall back to legacy single-date filenames."""
    current = snp_final_output(date_from, date_to)
    if current.exists():
        return current
    legacy = SNP_FINAL_DIR / f"CASRA_KPI_OUTPUT_{date_from}_FINAL.xlsx"
    if legacy.exists():
        return legacy
    flat_legacy = OUTPUT_ROOT / f"CASRA_KPI_OUTPUT_{date_from}_FINAL.xlsx"
    if flat_legacy.exists():
        return flat_legacy
    return current


def kpi_metrics_output(date_from: str, date_to: str) -> Path:
    suffix = output_period_suffix(date_from, date_to)
    return KPI_METRICS_DIR / f"CASRA_KPI_METRICS_{suffix}.xlsx"


def kpi_master_output() -> Path:
    return KPI_MASTER_DIR / "CASRA_KPI_METRICS_MASTER.xlsx"


def kpi_metrics_manual_output(run_date: str) -> Path:
    return KPI_METRICS_MANUAL_DIR / f"CASRA_KPI_METRICS_MANUAL_{run_date}.xlsx"


def hazmat_kpi_output(run_date: str) -> Path:
    return HAZMAT_KPI_DIR / f"CASRA_HAZMAT_KPI_{run_date}.xlsx"


# ============================================================================
# Excel helpers  (was casra_excel.py)
# ============================================================================
#
# Read/write helpers shared across the pipeline. These rely on the constants
# and date utilities defined above.


def validate_file(path: Path, label: str) -> Path:
    if not path.exists():
        raise FileNotFoundError(f"{label} file not found: {path}")
    return path


def normalize_header(value) -> str:
    text = str(value).replace("\u00a0", " ")
    return " ".join(text.strip().split())


def normalize_material_key(value) -> str:
    """Stable part-number key for matching ZMNM, Final Output, and DQ files."""
    if pd.isna(value):
        return ""
    if isinstance(value, bool):
        return str(value).strip().upper()
    if isinstance(value, int):
        return str(value).upper()
    if isinstance(value, float):
        if value == int(value):
            return str(int(value))
        text = str(value).strip()
    else:
        text = str(value).replace("\u00a0", " ").strip()
    if text.endswith(".0"):
        head = text[:-2]
        if head.isdigit() or (head.startswith("-") and head[1:].isdigit()):
            text = head
    return text.upper()


def find_col(df: pd.DataFrame, possible_names: list[str], label: str | None = None) -> str:
    lookup = {normalize_header(col).lower(): col for col in df.columns}
    for name in possible_names:
        key = normalize_header(name).lower()
        if key in lookup:
            return lookup[key]
    display = label or possible_names[0]
    raise KeyError(
        f"Missing {display} column. Tried: {possible_names}\n"
        f"Columns found in file: {list(df.columns)}"
    )


def read_excel_access(path: Path) -> pd.DataFrame:
    """Load workbook for access-db (string dtype, Access-style null handling)."""
    df = pd.read_excel(path, dtype=str)
    df.columns = [normalize_header(c) for c in df.columns]
    for col in df.columns:
        df[col] = df[col].astype("string").replace({"": pd.NA, "nan": pd.NA, "None": pd.NA})
    return df


def read_excel_table(path: Path, dtype=object) -> pd.DataFrame:
    df = pd.read_excel(path, dtype=dtype)
    df.columns = [normalize_header(c) for c in df.columns]
    return df


def add_report_date_column(df: pd.DataFrame, report_date: date | None = None) -> pd.DataFrame:
    """Insert Report Date as the first column on a detail output sheet."""
    df = df.copy()
    if REPORT_DATE_COL in df.columns:
        df = df.drop(columns=[REPORT_DATE_COL])
    df.insert(0, REPORT_DATE_COL, report_date or date.today())
    return df


def read_run_summary(path: Path) -> pd.DataFrame:
    try:
        return pd.read_excel(path, sheet_name="Run Summary", dtype=object)
    except (ValueError, KeyError):
        return pd.DataFrame()


def read_data_quality_file(path: Path) -> pd.DataFrame:
    """Load DQ workbook; auto-detect sheet and header row."""
    xl = pd.ExcelFile(path)

    for sheet_name in xl.sheet_names:
        raw = pd.read_excel(path, sheet_name=sheet_name, header=None, dtype=object)
        for header_row in range(min(20, len(raw))):
            headers = [normalize_header(v) for v in raw.iloc[header_row].tolist()]
            probe = pd.DataFrame(raw.iloc[header_row + 1 :].values, columns=headers)
            probe = probe.dropna(how="all").reset_index(drop=True)
            if probe.empty:
                continue
            try:
                find_col(probe, DQ_DATE_COLUMNS, "Date")
                find_col(probe, DQ_PART_COLUMNS, "P/N")
                print(f"  Data Quality: sheet '{sheet_name}', header row {header_row + 1}")
                return probe
            except KeyError:
                continue

    df = read_excel_table(path)
    find_col(df, DQ_DATE_COLUMNS, "Date")
    find_col(df, DQ_PART_COLUMNS, "P/N")
    return df


def period_from_run_summary(path: Path) -> tuple[str, str]:
    """Return (date_from, date_to) as YYYYMMDD strings from Run Summary."""
    summary = read_run_summary(path)
    if summary.empty:
        return "", ""

    def as_yyyymmdd(value) -> str:
        if pd.isna(value):
            return ""
        converted = yyyymmdd_to_date(value)
        if converted is not None:
            return converted.strftime("%Y%m%d")
        text = str(value).strip()
        if "." in text:
            text = text.split(".", 1)[0]
        return text[:8] if len(text) >= 8 and text[:8].isdigit() else ""

    date_from = as_yyyymmdd(summary["Date From"].iloc[0]) if "Date From" in summary.columns else ""
    date_to = as_yyyymmdd(summary["Date To"].iloc[0]) if "Date To" in summary.columns else ""
    return date_from, date_to


def coerce_check_columns(df: pd.DataFrame, columns: list[str]) -> pd.DataFrame:
    df = df.copy()
    for col in columns:
        if col not in df.columns:
            df[col] = 0
        df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0).astype(int)
    return df
