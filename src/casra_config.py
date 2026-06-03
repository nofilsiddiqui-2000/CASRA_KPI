"""Read config.txt next to the pipeline scripts (SAP credentials, etc.)."""

import os
from pathlib import Path


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
