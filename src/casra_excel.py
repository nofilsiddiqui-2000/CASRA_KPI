"""Shared Excel read/write helpers for the CASRA KPI pipeline."""

from __future__ import annotations

from pathlib import Path

import pandas as pd

from casra_constants import DQ_DATE_COLUMNS, DQ_PART_COLUMNS, REPORT_DATE_COL
from casra_dates import yyyymmdd_to_date


def validate_file(path: Path, label: str) -> Path:
    if not path.exists():
        raise FileNotFoundError(f"{label} file not found: {path}")
    return path


def normalize_header(value) -> str:
    text = str(value).replace("\u00a0", " ")
    return " ".join(text.strip().split())


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


RUN_SUMMARY_LEADING_COLUMNS = [REPORT_DATE_COL, "Date From", "Date To"]


def order_run_summary(df: pd.DataFrame) -> pd.DataFrame:
    """Put run metadata columns first for easier review."""
    if df.empty:
        return df
    leading = [c for c in RUN_SUMMARY_LEADING_COLUMNS if c in df.columns]
    rest = [c for c in df.columns if c not in leading]
    return df[leading + rest]


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
