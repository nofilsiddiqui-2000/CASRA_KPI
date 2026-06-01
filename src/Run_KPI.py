# Single entry point for the CASRA KPI pipeline.
#
# Prompts for one of two run modes:
#   1) Automated  - previous calendar month (no further input needed).
#   2) Manual     - user enters a custom start and end date once.
#
# The selected (date_from, date_to) is passed to every step via
# --date-from / --date-to so the SAP extracts, access-db output, and
# SNP exceptions all align on the same date range.
#
# Make sure to update the export paths in the SAP extract scripts to point
# to a shared location that this script can access, and that access-db.py
# can also access to pull the exported files from.

from datetime import datetime, timedelta
from pathlib import Path
import subprocess
import sys


BASE_DIR = Path(__file__).parent  # __file__ is the path to this script, so parent is the directory containing it.


SCRIPTS = [
    "Main_SAP_ZMNM_xl.py",
    "Main_SAP_ZMMR2199M_xl.py",
    "access-db.py",
    "apply_snp_exceptions.py",
]


def previous_month_range() -> tuple[str, str]:
    first_of_this_month = datetime.today().date().replace(day=1)
    last_of_prev_month = first_of_this_month - timedelta(days=1)
    first_of_prev_month = last_of_prev_month.replace(day=1)
    return (
        first_of_prev_month.strftime("%Y%m%d"),
        last_of_prev_month.strftime("%Y%m%d"),
    )


def prompt_yyyymmdd(prompt: str) -> str:
    while True:
        value = input(prompt).strip()
        try:
            datetime.strptime(value, "%Y%m%d")
            return value
        except ValueError:
            print("  Invalid date. Use YYYYMMDD format (e.g. 20260330).")


def prompt_manual_range() -> tuple[str, str]:
    while True:
        date_from = prompt_yyyymmdd("  Start date (YYYYMMDD): ")
        date_to = prompt_yyyymmdd("  End date   (YYYYMMDD): ")
        if date_from <= date_to:
            return date_from, date_to
        print("  Start date must be on or before end date. Try again.\n")


def select_mode() -> tuple[str, str]:
    print("\nCASRA KPI - Run Configuration")
    print("-" * 40)
    print("  1) Automated  -  Previous calendar month")
    print("  2) Manual     -  Enter a custom date range")
    while True:
        choice = input("\nSelect run mode [1/2]: ").strip()
        if choice == "1":
            df, dt = previous_month_range()
            print(f"\n  Using previous month: {df} -> {dt}")
            return df, dt
        if choice == "2":
            print()
            return prompt_manual_range()
        print("  Invalid selection. Enter 1 or 2.")


def run_script(script_name: str, date_from: str, date_to: str) -> None:
    script_path = BASE_DIR / script_name

    if not script_path.exists():
        raise FileNotFoundError(f"Script not found: {script_path}")

    print(f"\n>>> Running {script_name} ...")

    result = subprocess.run(
        [
            sys.executable, str(script_path),
            "--date-from", date_from,
            "--date-to", date_to,
        ],
        cwd=BASE_DIR,
    )

    if result.returncode != 0:
        raise RuntimeError(f"Error while running {script_name}. Stopping.")


def main() -> None:
    date_from, date_to = select_mode()

    for script in SCRIPTS:
        run_script(script, date_from, date_to)

    print("\nDone. SAP exports, KPI output, and SNP exceptions completed successfully.")
    print(f"Date range used: {date_from} -> {date_to}")


if __name__ == "__main__":
    main()
