"""
refresh.py
==========
Multi-country refresh.

    python refresh.py                # rebuild all country workbooks found
    python refresh.py argentina      # rebuild just Argentina
    python refresh.py argentina chile

What gets built per country
---------------------------
- Reads codes_<country>.csv on first run, or the Codes sheet of an existing
  Macro_Tracker_<Country>.xlsx on subsequent runs.
- Pulls every series via Haver DLX (or CSV / synthetic fallback).
- Rewrites Macro_Tracker_<Country>.xlsx with fresh data, forecasts, charts.

To add a third country: drop a codes_<name>.csv next to this file. Done.
"""

from __future__ import annotations

import sys
from pathlib import Path

from macro_tracker import (
    build_workbook,
    read_codes_csv,
    read_codes_from_workbook,
)

HERE = Path(__file__).parent


def _starter_csv(country_slug: str) -> Path:
    return HERE / f"codes_{country_slug}.csv"


def _workbook_path(country_label: str) -> Path:
    return HERE / f"Macro_Tracker_{country_label}.xlsx"


def _discover_countries() -> list[str]:
    """Find every codes_<name>.csv in this folder."""
    names = []
    for p in sorted(HERE.glob("codes_*.csv")):
        slug = p.stem.replace("codes_", "")
        names.append(slug)
    return names


def _label(slug: str) -> str:
    return slug.replace("_", " ").title()


def refresh_country(slug: str) -> None:
    label = _label(slug)
    starter = _starter_csv(slug)
    wb_path = _workbook_path(label)

    if wb_path.exists():
        codes = read_codes_from_workbook(wb_path)
        if codes is None or codes.empty:
            print(f"[note] Codes sheet empty in {wb_path.name}, falling back to {starter.name}")
            codes = read_codes_csv(starter)
        else:
            print(f"\n=== {label} ===  (read {len(codes)} codes from {wb_path.name})")
    else:
        if not starter.exists():
            print(f"[skip] no codes file for '{slug}' (looked for {starter.name})")
            return
        print(f"\n=== {label} ===  (bootstrapping from {starter.name})")
        codes = read_codes_csv(starter)

    build_workbook(codes, wb_path, country=label)


def main() -> None:
    args = sys.argv[1:]
    countries = args if args else _discover_countries()
    if not countries:
        print("No country codes files found. Create codes_<country>.csv to get started.")
        return
    for slug in countries:
        refresh_country(slug.lower())


if __name__ == "__main__":
    main()
