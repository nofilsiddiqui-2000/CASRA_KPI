# Run this main script. it will trigger the SAP export scripts, 
# and then the access-db script to process the exports. 
# Make sure to update the export paths in the SAP export scripts to point to a shared location 
# that this script can access, and that the access-db script can also access to pull the exported files from. 


from pathlib import Path
import subprocess
import sys


BASE_DIR = Path(__file__).parent # _file_ is the path to this script, so parent is the directory containing it. This should work in both testing and production environments without modification.



scripts = [
    "Main_SAP_ZMNM_xl.py",
    "Main_SAP_ZMMR2199M_xl.py",
    "access-db.py",
]

def run_script(script_name: str) -> None:
    script_path = BASE_DIR / script_name

    if not script_path.exists():
        raise FileNotFoundError(f"Script not found: {script_path}")

    print(f"\nRunning {script_name}...")

    result = subprocess.run(
        [sys.executable, str(script_path)],
        cwd=BASE_DIR
    )

    if result.returncode != 0:
        raise RuntimeError(f"Error while running {script_name}. Stopping.")


def main() -> None:
    for script in scripts:
        run_script(script)

    print("\nDone. SAP exports and KPI output completed successfully.")


if __name__ == "__main__":
    main()
