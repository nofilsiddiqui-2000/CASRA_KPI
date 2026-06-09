"""Shared folder layout for CASRA_KPI_OUTPUT.

Keeps each pipeline stage's files in its own subfolder (same idea as
SAP_Extracts/ZMNM and SAP_Extracts/ZMMR2199M).
"""

from __future__ import annotations

import shutil
from pathlib import Path


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

# generate_hazmat_kpi.py — standalone HazMat KPI (not yet in Run_KPI pipeline)
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
