"""Shared date utilities for the CASRA KPI pipeline.

Every step of the pipeline (SAP extracts, access-db, apply_snp_exceptions)
agrees on a single date range. Each script accepts:

    --date-from YYYYMMDD
    --date-to   YYYYMMDD

When omitted, both default to the previous calendar month. This way each
script still runs standalone, while Run_KPI.py can drive the whole chain
with one consistent range.
"""

from __future__ import annotations

import argparse
from datetime import date, datetime, timedelta


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
