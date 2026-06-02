"""Shared folder layout for CASRA_KPI_OUTPUT.

Keeps each pipeline stage's files in its own subfolder (same idea as
SAP_Extracts/ZMNM and SAP_Extracts/ZMMR2199M).
"""

from __future__ import annotations

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

# generate_kpi_metrics.py — accumulating history for Power BI
KPI_MASTER_DIR = OUTPUT_ROOT / "KPI_Master"

LOOKUP_DIR = ROOT_DIR / "LookUp Tables"
SAP_DIR = ROOT_DIR / "SAP_Extracts"


def ensure_output_dirs() -> None:
    for folder in (
        INTERMEDIATE_DIR,
        SNP_FINAL_DIR,
        KPI_METRICS_DIR,
        KPI_METRICS_MANUAL_DIR,
        KPI_MASTER_DIR,
    ):
        folder.mkdir(parents=True, exist_ok=True)


def intermediate_output(date_from: str) -> Path:
    return INTERMEDIATE_DIR / f"CASRA_KPI_OUTPUT_{date_from}.xlsx"


def snp_final_output(date_from: str) -> Path:
    return SNP_FINAL_DIR / f"CASRA_KPI_OUTPUT_{date_from}_FINAL.xlsx"


def kpi_metrics_output(date_from: str) -> Path:
    return KPI_METRICS_DIR / f"CASRA_KPI_METRICS_{date_from}.xlsx"


def kpi_master_output() -> Path:
    return KPI_MASTER_DIR / "CASRA_KPI_METRICS_MASTER.xlsx"


def kpi_metrics_manual_output(run_date: str) -> Path:
    return KPI_METRICS_MANUAL_DIR / f"CASRA_KPI_METRICS_MANUAL_{run_date}.xlsx"
