"""
EMBI Global Dashboard Builder
=============================

Reads any combination of J.P. Morgan EMBI Global / EMBI Global Diversified
CSVs and produces a single structured Excel workbook. The script auto-
detects three JPM file types and merges them:

  1. RETURNS file        — wide CSV, columns "EM Debt Indices | <Entity> | <Metric>"
                           Metrics: Cum Tot Ret Idx, Yld to Maturity, Z Spread,
                           Index Weight (%) (often empty / legacy EMBI only).
                           One row per date.

  2. WEIGHTS HISTORY     — two-row header. Row 1: FC_EMBIG_* codes.
                           Row 2: country names. Column 1: Trade Date,
                           column 2: Composite Index Weight, columns 3+:
                           per-country weight time series (back to 1993).

  3. SNAPSHOT            — flat CSV, one row per entity. Header includes
                           "Bam Id", "Instrument", "Mkt Cap %", "Average S&P
                           Rating", "Yield to Worst", "Z Spread to Worst",
                           "Spread Duration", and various return metrics
                           (Daily / MTD / YTD). Single trade date.

Multiple files of the same type are merged: when two files cover overlapping
dates, the later-modified file wins on conflict, but unique dates from older
files are preserved. This means once you have a deep history saved on disk,
re-running with just a fresh daily snapshot still produces a workbook with
the full history intact.

Usage
-----
    # Scan the current directory for *.csv files and process whatever it finds:
    python embi_builder.py

    # Or scan a specific directory:
    python embi_builder.py /path/to/EMBI_data

    # Or pass explicit files:
    python embi_builder.py returns.csv weights_history.csv snapshot.csv

    # Override the output filename:
    python embi_builder.py /path/to/data --output Dashboard_2026-05-07.xlsx
"""

from __future__ import annotations

import argparse
import csv
import math
import re
import statistics
import sys
from collections import defaultdict
from datetime import datetime
from pathlib import Path
from typing import Any, Dict, Iterable, List, Optional, Tuple

from openpyxl import Workbook
from openpyxl.chart import LineChart, BarChart, Reference
from openpyxl.formatting.rule import ColorScaleRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter


# ============================================================================
# 1. TAXONOMY — regions, rating order, name aliases
# ============================================================================

REGION_ORDER: List[str] = ["LatAm", "Europe", "Africa", "GCC", "Asia"]

COUNTRY_REGION: Dict[str, str] = {
    # LatAm
    "Argentina": "LatAm", "Barbados": "LatAm", "Belize": "LatAm",
    "Bolivia": "LatAm", "Brazil": "LatAm", "Chile": "LatAm",
    "Colombia": "LatAm", "Costa Rica": "LatAm", "Dominican Republic": "LatAm",
    "Ecuador": "LatAm", "El Salvador": "LatAm", "Guatemala": "LatAm",
    "Honduras": "LatAm", "Jamaica": "LatAm", "Mexico": "LatAm",
    "Panama": "LatAm", "Paraguay": "LatAm", "Peru": "LatAm",
    "Suriname": "LatAm", "Trinidad & Tobago": "LatAm",
    "Uruguay": "LatAm", "Venezuela": "LatAm",
    # Europe (Emerging Europe + Caucasus per JPM)
    "Armenia": "Europe", "Belarus": "Europe", "Bulgaria": "Europe",
    "Croatia": "Europe", "Georgia": "Europe", "Greece": "Europe",
    "Hungary": "Europe", "Latvia": "Europe", "Lithuania": "Europe",
    "Montenegro": "Europe", "Poland": "Europe", "Romania": "Europe",
    "Russia": "Europe", "Serbia": "Europe", "Slovak Republic": "Europe",
    "Turkey": "Europe", "Ukraine": "Europe",
    # Africa
    "Algeria": "Africa", "Angola": "Africa", "Benin": "Africa",
    "Cameroon": "Africa", "Cote d Ivoire": "Africa", "Egypt": "Africa",
    "Ethiopia": "Africa", "Gabon": "Africa", "Ghana": "Africa",
    "Kenya": "Africa", "Morocco": "Africa", "Mozambique": "Africa",
    "Namibia": "Africa", "Nigeria": "Africa", "Republic of Congo": "Africa",
    "Rwanda": "Africa", "Senegal": "Africa", "South Africa": "Africa",
    "Tanzania": "Africa", "Tunisia": "Africa", "Zambia": "Africa",
    # GCC / Mideast
    "Bahrain": "GCC", "Iraq": "GCC", "Jordan": "GCC", "Kuwait": "GCC",
    "Lebanon": "GCC", "Oman": "GCC", "Qatar": "GCC", "Saudi Arabia": "GCC",
    "UAE": "GCC",
    # Asia
    "Azerbaijan": "Asia", "China": "Asia", "India": "Asia",
    "Indonesia": "Asia", "Kazakhstan": "Asia", "Kyrgyzstan": "Asia",
    "Malaysia": "Asia", "Maldives": "Asia", "Mongolia": "Asia",
    "Pakistan": "Asia", "Papua New Guinea": "Asia", "Philippines": "Asia",
    "South Korea": "Asia", "Sri Lanka": "Asia", "Tajikistan": "Asia",
    "Thailand": "Asia", "Uzbekistan": "Asia", "Vietnam": "Asia",
}

REGION_AGGREGATE: Dict[str, str] = {
    "LatAm": "Latin Region",
    "Europe": "Europe Region",
    "Africa": "Africa Region",
    "GCC": "Mideast Region",
    "Asia": "Asia Region",
}

# JPM "By Region" labels in the snapshot file map differently.
SNAPSHOT_REGION_LABEL: Dict[str, str] = {
    "Africa": "Africa", "Asia": "Asia", "Europe": "Europe",
    "Latin": "LatAm", "Middle East": "GCC",
}

RATING_ORDER: List[str] = [
    "Credit AA only", "Credit A only", "Credit BBB only",
    "Credit BB only", "Credit B only", "Credit C only",
    "Credit IG only", "Credit Non-IG", "Credit NR", "Credit Residual only",
]
RATING_DISPLAY: Dict[str, str] = {
    "Credit AA only": "AA",
    "Credit A only": "A",
    "Credit BBB only": "BBB",
    "Credit BB only": "BB",
    "Credit B only": "B",
    "Credit C only": "C / Distressed",
    "Credit IG only": "IG (composite)",
    "Credit Non-IG": "Non-IG (composite)",
    "Credit NR": "Not Rated",
    "Credit Residual only": "Residual",
}

# S&P / Moody-style rating to numeric score (higher = better quality).
# Used to sort countries by credit quality within a region.
RATING_SCORE: Dict[str, int] = {
    "AAA": 22, "Aaa": 22,
    "AA+": 21, "Aa1": 21,
    "AA": 20,  "Aa2": 20,
    "AA-": 19, "Aa3": 19,
    "A+": 18,  "A1": 18,
    "A": 17,   "A2": 17,
    "A-": 16,  "A3": 16,
    "BBB+": 15, "Baa1": 15,
    "BBB": 14, "Baa2": 14,
    "BBB-": 13, "Baa3": 13,
    "BB+": 12, "Ba1": 12,
    "BB": 11,  "Ba2": 11,
    "BB-": 10, "Ba3": 10,
    "B+": 9,   "B1": 9,
    "B": 8,    "B2": 8,
    "B-": 7,   "B3": 7,
    "CCC+": 6, "Caa1": 6,
    "CCC": 5,  "Caa2": 5,
    "CCC-": 4, "Caa3": 4,
    "CC": 3,   "Ca": 3,
    "C": 2,
    "D": 1, "SD": 1,
    "NR": 0,
}

LATAM_FOCUS: List[str] = [
    "Argentina", "Chile", "Colombia", "Dominican Republic",
    "Mexico", "Panama", "Peru", "Venezuela",
]

INDEX_NAME = "EMBI Global"
METRICS = {
    "spread": "Z Spread",
    "yield":  "Yld to Maturity",
    "tret":   "Cum Tot Ret Idx",
}

# Maps spellings used by snapshot / weights / returns files to a single canonical form.
NAME_ALIASES: Dict[str, str] = {
    "Cote D'Ivoire":         "Cote d Ivoire",
    "Cote d'Ivoire":         "Cote d Ivoire",
    "Côte d'Ivoire":         "Cote d Ivoire",
    "Trinidad And Tobago":   "Trinidad & Tobago",
    "Trinidad and Tobago":   "Trinidad & Tobago",
    "Trinidad &amp; Tobago": "Trinidad & Tobago",
    "Slovakia":              "Slovak Republic",
    "Republic of the Congo": "Republic of Congo",
    "Republic Of Congo":     "Republic of Congo",
    "Congo, Republic of":    "Republic of Congo",
}


def normalize_name(name: str) -> str:
    name = (name or "").strip().replace("&amp;", "&").replace("&#38;", "&")
    return NAME_ALIASES.get(name, name)


# ============================================================================
# 2. STYLING
# ============================================================================

FONT_NAME = "Arial"
COLOR_HEADER_BG = "1F4E78"
COLOR_HEADER_FG = "FFFFFF"
COLOR_REGION_BG = "D9E1F2"
COLOR_RATING_BG = "FFF2CC"
COLOR_INDEX_BG  = "C6E0B4"
COLOR_INPUT_BG  = "FFFF00"   # bright yellow — flag user-editable assumption cells
COLOR_BORDER    = "BFBFBF"
COLOR_HARDCODE  = "0000FF"
COLOR_FORMULA   = "000000"
COLOR_CROSSREF  = "008000"
COLOR_NOTE      = "7F7F7F"

THIN_BORDER = Border(
    left=Side(style="thin", color=COLOR_BORDER),
    right=Side(style="thin", color=COLOR_BORDER),
    top=Side(style="thin", color=COLOR_BORDER),
    bottom=Side(style="thin", color=COLOR_BORDER),
)


# ============================================================================
# 3. PARSERS
# ============================================================================

def _parse_date(raw: str) -> datetime:
    raw = (raw or "").strip()
    for fmt in ("%d-%b-%y", "%d-%b-%Y", "%Y-%m-%d", "%m/%d/%Y", "%d/%m/%Y"):
        try:
            return datetime.strptime(raw, fmt)
        except ValueError:
            continue
    raise ValueError(f"unrecognised date: {raw!r}")


def _parse_value(raw: str) -> Optional[float]:
    if raw is None:
        return None
    s = raw.strip()
    if s == "" or s.upper() in {"N/A", "NA", "#N/A", "NULL"}:
        return None
    try:
        return float(s.replace(",", ""))
    except ValueError:
        return None


def _open_csv(path: Path):
    """Open a JPM CSV with BOM-stripping and replace any decoding errors.

    Use this everywhere instead of plain open() so loaders are tolerant of
    file-format quirks (BOM, occasional non-UTF-8 bytes from JPM exports).
    """
    return path.open("r", newline="", encoding="utf-8-sig", errors="replace")


def _read_skipping_blank_rows(reader, n: int) -> List[List[str]]:
    """Return the next `n` non-empty rows from a csv.reader."""
    out: List[List[str]] = []
    for row in reader:
        if any((c or "").strip() for c in row):
            out.append(row)
            if len(out) >= n:
                return out
    return out


def classify_csv(path: Path) -> str:
    """Return one of 'returns', 'weights_history', 'snapshot', 'unknown'.

    Classification is based purely on CONTENT (header signatures), never on
    the filename. JPM ships files with timestamps, GUIDs, and inconsistent
    naming — that doesn't matter here. The script tolerates: UTF-8 BOM,
    leading blank rows, header whitespace, case differences, and quoted
    cells.
    """
    try:
        # utf-8-sig silently strips the BOM if one is present.
        with path.open("r", newline="", encoding="utf-8-sig", errors="replace") as f:
            reader = csv.reader(f)
            # Skip any number of leading blank or all-empty-cell rows until
            # we find one that actually contains content.
            row1: List[str] = []
            for candidate in reader:
                if any((c or "").strip() for c in candidate):
                    row1 = candidate
                    break
            if not row1:
                return "unknown"
            # Same for row 2 — skip blank rows there too.
            row2: List[str] = []
            for candidate in reader:
                if any((c or "").strip() for c in candidate):
                    row2 = candidate
                    break
    except (OSError, StopIteration):
        return "unknown"

    def _norm(s):
        return (s or "").strip().strip('"').strip().lower()

    head = _norm(row1[0])

    # ---- RETURNS: first column "Date", subsequent columns contain
    #      "EM Debt Indices | <Entity> | <Metric>".
    if head == "date" and any("em debt indices" in (c or "").lower() for c in row1[1:]):
        return "returns"

    # ---- SNAPSHOT: first column "Bam Id" (any spacing/capitalization) and
    #      "Instrument" appearing as another column header.
    if head in ("bam id", "bamid", "bam_id"):
        if any(_norm(c) == "instrument" for c in row1):
            return "snapshot"
        # Even without "Instrument", the Bam Id signature alone is rare enough
        # to be conclusive — JPM uses it specifically for these dumps.
        return "snapshot"

    # ---- WEIGHTS HISTORY: two-row header. Row 1 has FC_EMBIG_* codes.
    #      Row 2 starts with "Trade Date" and has country names from col 3+.
    has_codes = any((c or "").strip().upper().startswith("FC_EMBIG") for c in row1)
    has_trade_date = bool(row2) and _norm(row2[0]) == "trade date"
    if has_codes and has_trade_date:
        return "weights_history"
    # Fallback: if row 1 is empty in col 0 and row 2 says "Trade Date" with a
    # "Composite Index Weight" column, we're still looking at a weights file
    # — don't insist on FC_EMBIG codes (JPM might rename them in a future build).
    if has_trade_date and any("composite" in (c or "").lower() and "weight" in (c or "").lower() for c in row2):
        return "weights_history"

    return "unknown"


def load_returns(path: Path) -> Tuple[List[datetime], Dict[Tuple[str, str], List[Optional[float]]]]:
    """Load a returns/yields/spreads file. Returns ([dates], {(entity,metric): [values]})."""
    with _open_csv(path) as f:
        reader = csv.reader(f)
        first = _read_skipping_blank_rows(reader, 1)
        header = first[0] if first else []
        rows = list(reader)

    columns: List[Tuple[str, str]] = []
    for col in header[1:]:
        cleaned = (col or "").replace("&amp;", "&").replace("&#38;", "&")
        parts = [p.strip() for p in cleaned.split("|")]
        if len(parts) < 3:
            columns.append(("", ""))
            continue
        columns.append((normalize_name(parts[1]), parts[2]))

    parsed: List[Tuple[datetime, List[Optional[float]]]] = []
    for row in rows:
        if not row or not row[0].strip():
            continue
        try:
            d = _parse_date(row[0])
        except ValueError:
            continue
        parsed.append((d, [_parse_value(c) for c in row[1:]]))
    parsed.sort(key=lambda r: r[0])
    dates = [r[0] for r in parsed]

    series: Dict[Tuple[str, str], List[Optional[float]]] = {}
    for idx, key in enumerate(columns):
        if key == ("", ""):
            continue
        series.setdefault(key, [None] * len(parsed))
        for i, (_, vals) in enumerate(parsed):
            if idx < len(vals):
                series[key][i] = vals[idx]
    return dates, series


def load_weights_history(path: Path) -> Tuple[List[datetime], Dict[str, List[Optional[float]]]]:
    """Load monthly weights history (two-row header). Returns ([dates], {country: [weight]})."""
    with _open_csv(path) as f:
        reader = csv.reader(f)
        first_two = _read_skipping_blank_rows(reader, 2)
        codes = first_two[0] if first_two else []
        names = first_two[1] if len(first_two) > 1 else []
        rows = list(reader)

    countries = [normalize_name(n) for n in names[2:]]
    parsed: List[Tuple[datetime, List[Optional[float]]]] = []
    for row in rows:
        if not row or not row[0].strip():
            continue
        try:
            d = _parse_date(row[0])
        except ValueError:
            continue
        # Skip composite weight column 1, take per-country from column 2 onwards.
        parsed.append((d, [_parse_value(c) for c in row[2:]]))
    parsed.sort(key=lambda r: r[0])
    dates = [r[0] for r in parsed]
    out: Dict[str, List[Optional[float]]] = {c: [None] * len(parsed) for c in countries}
    for i, (_, vals) in enumerate(parsed):
        for j, c in enumerate(countries):
            if j < len(vals):
                out[c][i] = vals[j]
    return dates, out


def load_snapshot(path: Path) -> Tuple[Optional[datetime], Dict[str, Dict[str, Any]]]:
    """Load a single-date snapshot file. Returns (snapshot_date, {entity: {field: val}})."""
    with _open_csv(path) as f:
        reader = csv.DictReader(f)
        rows = list(reader)
    snap_date: Optional[datetime] = None
    out: Dict[str, Dict[str, Any]] = {}
    for row in rows:
        instrument = (row.get("Instrument") or "").strip()
        if not instrument or instrument.startswith("By "):
            continue
        if snap_date is None:
            try:
                snap_date = _parse_date(row.get("Date", ""))
            except (ValueError, KeyError):
                pass
        # Snapshot uses "Latin", "Middle East", etc. — translate to our region keys.
        if instrument in SNAPSHOT_REGION_LABEL:
            key = f"REGION:{SNAPSHOT_REGION_LABEL[instrument]}"
        elif instrument == "Non Latin":
            key = "AGG:NonLatin"
        elif instrument == "EMBI Global Diversified":
            key = "INDEX:EMBIGD"
        else:
            key = normalize_name(instrument)
        out[key] = row
    return snap_date, out


# ============================================================================
# 4. AGGREGATOR — merge multiple files of the same type
# ============================================================================

def merge_returns(file_results: List[Tuple[List[datetime], Dict[Tuple[str, str], List[Optional[float]]]]]
                  ) -> Tuple[List[datetime], Dict[Tuple[str, str], List[Optional[float]]]]:
    if not file_results:
        return [], {}
    # Build a (key) -> {date -> value} dict, then collapse to sorted list.
    by_key: Dict[Tuple[str, str], Dict[datetime, float]] = defaultdict(dict)
    all_dates: set = set()
    for dates, series in file_results:
        for d in dates:
            all_dates.add(d)
        for key, vals in series.items():
            for i, d in enumerate(dates):
                if vals[i] is not None:
                    by_key[key][d] = vals[i]
    sorted_dates = sorted(all_dates)
    out_series: Dict[Tuple[str, str], List[Optional[float]]] = {}
    for key, dmap in by_key.items():
        out_series[key] = [dmap.get(d) for d in sorted_dates]
    return sorted_dates, out_series


def merge_weights_history(file_results: List[Tuple[List[datetime], Dict[str, List[Optional[float]]]]]
                          ) -> Tuple[List[datetime], Dict[str, List[Optional[float]]]]:
    if not file_results:
        return [], {}
    by_key: Dict[str, Dict[datetime, float]] = defaultdict(dict)
    all_dates: set = set()
    for dates, series in file_results:
        for d in dates:
            all_dates.add(d)
        for key, vals in series.items():
            for i, d in enumerate(dates):
                if vals[i] is not None:
                    by_key[key][d] = vals[i]
    sorted_dates = sorted(all_dates)
    out: Dict[str, List[Optional[float]]] = {}
    for key, dmap in by_key.items():
        out[key] = [dmap.get(d) for d in sorted_dates]
    return sorted_dates, out


def merge_snapshots(file_results: List[Tuple[Optional[datetime], Dict[str, Dict[str, Any]]]]
                    ) -> Tuple[Optional[datetime], Dict[str, Dict[str, Any]]]:
    """Use the latest snapshot file as the active one (older snapshots discarded)."""
    if not file_results:
        return None, {}
    file_results = sorted(
        [(d, s) for d, s in file_results if d is not None],
        key=lambda r: r[0],
    )
    if not file_results:
        return None, {}
    return file_results[-1]


# ============================================================================
# 5. BUILDER
# ============================================================================

class Builder:
    def __init__(
        self,
        dates: List[datetime],
        series: Dict[Tuple[str, str], List[Optional[float]]],
        weight_dates: List[datetime],
        weight_series: Dict[str, List[Optional[float]]],
        snap_date: Optional[datetime],
        snap_data: Dict[str, Dict[str, Any]],
        sources: List[Tuple[str, str]],  # [(file_kind, filename)]
    ) -> None:
        self.dates = dates
        self.series = series
        self.weight_dates = weight_dates
        self.weight_series = weight_series
        self.snap_date = snap_date
        self.snap_data = snap_data
        self.sources = sources

        self.frequency = self._detect_frequency(dates)
        self.wb = Workbook()
        self.wb.remove(self.wb.active)

        # Resolve countries actually present in returns data.
        present = {ent for (ent, _) in self.series.keys()}
        self.countries_by_region: Dict[str, List[str]] = {r: [] for r in REGION_ORDER}
        for entity, region in COUNTRY_REGION.items():
            if entity in present or entity in self.weight_series or entity in self.snap_data:
                self.countries_by_region[region].append(entity)

        # Sort countries within each region by S&P rating score (best first), then alphabetically.
        for region in REGION_ORDER:
            self.countries_by_region[region].sort(
                key=lambda c: (-self._rating_score(c), c)
            )

        self.all_countries: List[str] = [
            c for r in REGION_ORDER for c in self.countries_by_region[r]
        ]
        self.ratings_present: List[str] = [r for r in RATING_ORDER if r in present]

    # --------- helpers ----------
    @staticmethod
    def _detect_frequency(dates: List[datetime]) -> str:
        if len(dates) < 2:
            return "unknown"
        deltas = [(dates[i] - dates[i - 1]).days for i in range(1, len(dates))]
        avg = sum(deltas) / len(deltas)
        if avg <= 3:   return "daily"
        if 6 <= avg <= 9: return "weekly"
        if 25 <= avg <= 35: return "monthly"
        return "mixed"

    def _rating_score(self, country: str) -> int:
        snap = self.snap_data.get(country)
        if snap:
            for field in ("Average S&P Rating", "Average Moody Rating", "Average Fitch Rating"):
                v = (snap.get(field) or "").strip()
                if v in RATING_SCORE:
                    return RATING_SCORE[v]
        return -1  # unknown ratings sort to the bottom of their region

    def latest_value(self, entity: str, metric: str) -> Optional[float]:
        vals = self.series.get((entity, metric))
        if not vals:
            return None
        for v in reversed(vals):
            if v is not None:
                return v
        return None

    def value_at(self, entity: str, metric: str, idx: int) -> Optional[float]:
        vals = self.series.get((entity, metric))
        if not vals or idx is None or idx < 0 or idx >= len(vals):
            return None
        return vals[idx]

    def year_start_index(self, year: int) -> Optional[int]:
        if not self.dates:
            return None
        prior = None
        for i, d in enumerate(self.dates):
            if d.year < year:
                prior = i
            elif d.year >= year:
                break
        return prior if prior is not None else 0

    def latest_weight(self, country: str) -> Optional[float]:
        vals = self.weight_series.get(country)
        if not vals:
            return None
        for v in reversed(vals):
            if v is not None:
                return v
        return None

    def latest_weight_date(self) -> Optional[datetime]:
        if not self.weight_dates:
            return None
        # find the most recent date with at least one non-null weight
        for i in range(len(self.weight_dates) - 1, -1, -1):
            for c in self.weight_series:
                if self.weight_series[c][i] is not None:
                    return self.weight_dates[i]
        return None

    # --------- cell helpers ----------
    def _font(self, cell, **kw):
        d = {"name": FONT_NAME, "size": 10}
        d.update(kw)
        cell.font = Font(**d)

    def _hdr(self, cell, txt):
        cell.value = txt
        cell.fill = PatternFill("solid", start_color=COLOR_HEADER_BG)
        self._font(cell, color=COLOR_HEADER_FG, bold=True)
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = THIN_BORDER

    def _region_lbl(self, cell, txt):
        cell.value = txt
        cell.fill = PatternFill("solid", start_color=COLOR_REGION_BG)
        self._font(cell, bold=True)
        cell.alignment = Alignment(horizontal="left", vertical="center")
        cell.border = THIN_BORDER

    def _rating_lbl(self, cell, txt):
        cell.value = txt
        cell.fill = PatternFill("solid", start_color=COLOR_RATING_BG)
        self._font(cell, bold=True, italic=True)
        cell.alignment = Alignment(horizontal="left", vertical="center")
        cell.border = THIN_BORDER

    def _index_lbl(self, cell, txt):
        cell.value = txt
        cell.fill = PatternFill("solid", start_color=COLOR_INDEX_BG)
        self._font(cell, bold=True)
        cell.alignment = Alignment(horizontal="left", vertical="center")
        cell.border = THIN_BORDER

    def _val(self, cell, value, *, fmt="0.00", color=COLOR_FORMULA):
        cell.value = value
        cell.number_format = fmt
        self._font(cell, color=color)
        cell.alignment = Alignment(horizontal="right", vertical="center")
        cell.border = THIN_BORDER

    # --------- Cover ----------
    def build_cover(self):
        ws = self.wb.create_sheet("Cover")
        ws.sheet_view.showGridLines = False
        ws.column_dimensions["A"].width = 30
        ws.column_dimensions["B"].width = 80

        ws["A1"] = "EMBI Global Dashboard"
        self._font(ws["A1"], size=20, bold=True, color=COLOR_HEADER_BG)
        ws.row_dimensions[1].height = 28

        meta = [
            ("Built on", datetime.now().strftime("%Y-%m-%d %H:%M")),
            ("YTD anchor year", str(datetime.now().year)),
            ("Returns frequency", self.frequency),
            ("Returns observations", len(self.dates)),
            ("Returns date range", f"{self.dates[0]:%Y-%m-%d} — {self.dates[-1]:%Y-%m-%d}" if self.dates else "(no returns data)"),
            ("Weights observations", len(self.weight_dates)),
            ("Weights date range", f"{self.weight_dates[0]:%Y-%m-%d} — {self.weight_dates[-1]:%Y-%m-%d}" if self.weight_dates else "(no weights data)"),
            ("Snapshot date", self.snap_date.strftime("%Y-%m-%d") if self.snap_date else "(no snapshot loaded)"),
            ("Countries present", sum(len(v) for v in self.countries_by_region.values())),
            ("Rating buckets present", len(self.ratings_present)),
        ]
        for i, (k, v) in enumerate(meta, start=3):
            ws.cell(row=i, column=1, value=k)
            self._font(ws.cell(row=i, column=1), bold=True)
            ws.cell(row=i, column=2, value=v)
            self._font(ws.cell(row=i, column=2))

        ws.cell(row=14, column=1, value="Source files loaded")
        self._font(ws.cell(row=14, column=1), size=12, bold=True, color=COLOR_HEADER_BG)
        for i, (kind, fname) in enumerate(self.sources, start=15):
            ws.cell(row=i, column=1, value=f"  • {kind}")
            self._font(ws.cell(row=i, column=1), italic=True)
            ws.cell(row=i, column=2, value=fname)
            self._font(ws.cell(row=i, column=2))

        legend_row = 15 + len(self.sources) + 2
        ws.cell(row=legend_row, column=1, value="Color key")
        self._font(ws.cell(row=legend_row, column=1), size=12, bold=True, color=COLOR_HEADER_BG)
        legend = [
            ("Hardcoded input (blue text)",   COLOR_HARDCODE),
            ("Formula / calculation (black)", COLOR_FORMULA),
            ("Cross-sheet link (green)",      COLOR_CROSSREF),
            ("Note / metadata (grey)",        COLOR_NOTE),
        ]
        for i, (lbl, color) in enumerate(legend, start=legend_row + 1):
            ws.cell(row=i, column=1, value=lbl)
            self._font(ws.cell(row=i, column=1), color=color, bold=True)

    # --------- Instructions ----------
    def build_instructions(self):
        ws = self.wb.create_sheet("Instructions")
        ws.sheet_view.showGridLines = False
        ws.column_dimensions["A"].width = 30
        ws.column_dimensions["B"].width = 100
        ws["A1"] = "How to keep this workbook up to date"
        self._font(ws["A1"], size=18, bold=True, color=COLOR_HEADER_BG)
        ws.row_dimensions[1].height = 26

        sections = [
            ("Workflow in one line",
             "Drop fresh JPM CSVs into your project folder, open a terminal in that folder, run "
             "`python embi_builder.py` — done. The workbook regenerates with all the data the script "
             "can find."),
            ("",""),
            ("Required files (download from JPM)",
             "Three file types power this workbook. You don't need all three every time, but you do "
             "need each at least once. The script auto-detects file type by inspecting HEADERS, not "
             "the filename — so JPM can rename, timestamp, GUID-suffix, or otherwise mangle the "
             "file name however it wants. As long as the column structure is intact, classification "
             "works. The script also tolerates UTF-8 BOM, leading blank rows, header whitespace, "
             "and case differences."),
            ("  1. RETURNS file",
             "JPM 'Markets / DataQuery' EMBI Global query. Columns named 'EM Debt Indices | <Entity> | "
             "<Metric>'. Metrics: Cum Tot Ret Idx, Yld to Maturity, Z Spread. One row per date. "
             "Typical filename: 'Query 3_<id>.csv'. Download whatever frequency you want — monthly to "
             "build history, then daily going forward. Both work."),
            ("  2. WEIGHTS HISTORY file",
             "JPM 'EMBI Global Diversified — Monthly' download. Two-row header: row 1 has FC_EMBIG_* "
             "ticker codes, row 2 has country names; column 1 is 'Trade Date'. Goes back to 1993. "
             "Typical filename: 'JPM_EMBI_Global_Div___Mo_<date>_<id>.csv'. Re-download whenever you "
             "want fresh weights — monthly is plenty."),
            ("  3. SNAPSHOT file",
             "JPM 'EMBI Global Diversified' single-date detail. Flat CSV starting with 'Bam Id'. "
             "Includes ratings (S&P/Moody/Fitch), spread duration, market cap, daily/MTD/YTD returns. "
             "Typical filename: 'JPM_EMBI_Global_Diversif_<date>_<id>.csv'. Refresh as often as you "
             "want; only the most recent snapshot is shown but earlier files are harmless."),
            ("","",),
            ("Adding new data",
             "Just drop the new file in the folder alongside the older ones and re-run the script. "
             "If two files cover the same dates the most recent file's data wins, but unique dates "
             "from older files are preserved — so your history is never lost."),
            ("","",),
            ("YTD calculation",
             "All YTD figures (returns, spread changes, yield changes) are anchored to the CURRENT "
             "calendar year (today's year, per your computer's clock) — not to the latest data point. "
             "If you re-run the script in February 2027 against data that ends Dec-2026, YTD will "
             "show as zero or near-zero (because we're at the start of the new year and have no 2027 "
             "data yet). The YTD anchor year is shown on the Cover tab so you can sanity-check it."),
            ("","",),
            ("Optional CLI flags",
             "`python embi_builder.py /path/to/folder` — point at a different folder.  "
             "`python embi_builder.py file1.csv file2.csv` — explicit list.  "
             "`python embi_builder.py --output My_Dashboard.xlsx` — change output filename.  "
             "`python embi_builder.py --recursive` — scan subfolders too."),
            ("","",),
            ("What each tab does",
             "Cover (metadata) → Instructions (this guide) → Forecast (interactive 3/6/12-month "
             "monitor — yellow cells are your assumptions) → Weights (latest snapshot, region-"
             "organized) → Spreads / Yields / TR_YTD (time series, country sorted by S&P rating "
             "within region) → By_Rating (composition by JPM rating bucket) → Snapshot (per-country "
             "deep dive: ratings, duration, mkt cap, returns) → LatAm_Focus (8 focus credits vs "
             "peers) → Charts (region-level overviews) → Weights_History (full weight time series) "
             "→ Data_Raw (long-form pivot of all loaded data)."),
            ("","",),
            ("If something looks wrong",
             "Check the 'Source files loaded' section on the Cover tab — it lists every file the "
             "script ingested, classified by type. If a file you expected isn't there, your filename "
             "may have been excluded by the glob (only *.csv) or the file format may have drifted "
             "from what JPM gives today. The script's classify_csv() function is the place to teach "
             "it about a new variant."),
            ("","",),
            ("Country-name normalization",
             "JPM uses different spellings for the same country across files (e.g. 'Cote d Ivoire' "
             "vs \"Cote D'Ivoire\", 'Trinidad & Tobago' vs 'Trinidad And Tobago'). The script "
             "normalizes these. If you see a country missing weights or rating data, search "
             "embi_builder.py for NAME_ALIASES and add the new variant."),
            ("","",),
            ("When in doubt",
             "The script regenerates the file in full each run. Nothing the user enters into the "
             "workbook is preserved across runs — so if you add notes or formulas, do it in a copy. "
             "All historical data is preserved because it lives in the source CSVs, not in the xlsx."),
        ]
        row = 3
        for k, v in sections:
            if not k and not v:
                row += 1
                continue
            ws.cell(row=row, column=1, value=k)
            self._font(ws.cell(row=row, column=1), bold=True if k and not k.startswith("  ") else False, size=11 if not k.startswith("  ") else 10)
            ws.cell(row=row, column=1).alignment = Alignment(vertical="top", wrap_text=True)
            ws.cell(row=row, column=2, value=v)
            self._font(ws.cell(row=row, column=2))
            ws.cell(row=row, column=2).alignment = Alignment(vertical="top", wrap_text=True)
            ws.row_dimensions[row].height = max(18, 14 * (1 + len(v) // 90))
            row += 1

    # --------- Weights ----------
    def build_weights(self):
        ws = self.wb.create_sheet("Weights")
        ws.sheet_view.showGridLines = False
        ws["A1"] = "EMBI Global — Country Weights (latest)"
        self._font(ws["A1"], size=14, bold=True, color=COLOR_HEADER_BG)
        ws.row_dimensions[1].height = 22

        latest_w_date = self.latest_weight_date()
        if latest_w_date:
            ws["A2"] = f"Source: weights history file. Latest date: {latest_w_date:%Y-%m-%d}. " \
                       f"All values pulled from the Weights_History tab via INDEX/MATCH."
        else:
            ws["A2"] = "No weights history loaded — paste a JPM weights file or update Weights_History tab manually."
        ws.merge_cells("A2:G2")
        self._font(ws["A2"], color=COLOR_NOTE, italic=True)
        ws["A2"].alignment = Alignment(wrap_text=True, vertical="top")
        ws.row_dimensions[2].height = 32

        header_row = 4
        cols = [
            ("Region", 14), ("Country", 24), ("S&P", 8), ("Moody's", 9), ("Fitch", 8),
            ("Weight (%)", 13), ("Region share (%)", 16),
            ("Latest Spread (bps)", 18), ("Latest Yield (%)", 16),
        ]
        for i, (h, w) in enumerate(cols, start=1):
            self._hdr(ws.cell(row=header_row, column=i), h)
            ws.column_dimensions[get_column_letter(i)].width = w

        row = header_row + 1
        region_subtotal_rows: Dict[str, int] = {}

        for region in REGION_ORDER:
            countries = self.countries_by_region.get(region, [])
            if not countries:
                continue
            self._region_lbl(ws.cell(row=row, column=1), region)
            self._region_lbl(ws.cell(row=row, column=2), f"{region} subtotal")
            for c in (3, 4, 5):
                ws.cell(row=row, column=c, value="")
                ws.cell(row=row, column=c).fill = PatternFill("solid", start_color=COLOR_REGION_BG)
                ws.cell(row=row, column=c).border = THIN_BORDER
            subtotal_row = row
            row += 1
            child_start = row
            for c in countries:
                ws.cell(row=row, column=1, value="").border = THIN_BORDER
                self._font(ws.cell(row=row, column=1))
                name_cell = ws.cell(row=row, column=2, value=c)
                self._font(name_cell)
                name_cell.border = THIN_BORDER

                # Ratings from snapshot
                snap = self.snap_data.get(c, {}) or {}
                self._val(ws.cell(row=row, column=3), (snap.get("Average S&P Rating") or "").strip() or "", fmt="@", color=COLOR_HARDCODE)
                self._val(ws.cell(row=row, column=4), (snap.get("Average Moody Rating") or "").strip() or "", fmt="@", color=COLOR_HARDCODE)
                self._val(ws.cell(row=row, column=5), (snap.get("Average Fitch Rating") or "").strip() or "", fmt="@", color=COLOR_HARDCODE)

                w = self.latest_weight(c)
                if w is None:
                    self._val(ws.cell(row=row, column=6), "", fmt="0.00%;(0.00%);-", color=COLOR_HARDCODE)
                else:
                    # Weights from JPM are stored as percentages (e.g. 5.12 = 5.12%); divide.
                    self._val(ws.cell(row=row, column=6), w / 100, fmt="0.00%;(0.00%);-", color=COLOR_HARDCODE)

                # Region share = country weight / region subtotal
                share = ws.cell(row=row, column=7)
                share.value = f'=IFERROR(F{row}/$F${subtotal_row},"")'
                share.number_format = "0.0%;(0.0%);-"
                self._font(share)
                share.alignment = Alignment(horizontal="right")
                share.border = THIN_BORDER

                sp = self.latest_value(c, METRICS["spread"])
                yd = self.latest_value(c, METRICS["yield"])
                self._val(ws.cell(row=row, column=8), sp if sp is not None else "", fmt="0", color=COLOR_HARDCODE)
                self._val(ws.cell(row=row, column=9), yd if yd is not None else "", fmt="0.00", color=COLOR_HARDCODE)
                row += 1
            child_end = row - 1
            sub_cell = ws.cell(row=subtotal_row, column=6)
            sub_cell.value = f"=IFERROR(SUM(F{child_start}:F{child_end}),0)"
            sub_cell.number_format = "0.00%;(0.00%);-"
            self._font(sub_cell, bold=True)
            sub_cell.alignment = Alignment(horizontal="right")
            sub_cell.border = THIN_BORDER
            sub_cell.fill = PatternFill("solid", start_color=COLOR_REGION_BG)
            region_subtotal_rows[region] = subtotal_row
            row += 1

        if region_subtotal_rows:
            self._index_lbl(ws.cell(row=row, column=1), "Total")
            self._index_lbl(ws.cell(row=row, column=2), "EMBI Global")
            for c in (3, 4, 5):
                ws.cell(row=row, column=c, value="")
                ws.cell(row=row, column=c).fill = PatternFill("solid", start_color=COLOR_INDEX_BG)
                ws.cell(row=row, column=c).border = THIN_BORDER
            sum_parts = "+".join(f"F{r}" for r in region_subtotal_rows.values())
            tot = ws.cell(row=row, column=6)
            tot.value = f"={sum_parts}" if sum_parts else "=0"
            tot.number_format = "0.00%;(0.00%);-"
            self._font(tot, bold=True)
            tot.alignment = Alignment(horizontal="right")
            tot.border = THIN_BORDER
            tot.fill = PatternFill("solid", start_color=COLOR_INDEX_BG)

        ws.freeze_panes = "A5"

        # Color scale on weights
        if region_subtotal_rows:
            ws.conditional_formatting.add(
                f"F{header_row + 1}:F{row}",
                ColorScaleRule(
                    start_type="min", start_color="FFFFFF",
                    end_type="max", end_color="63BE7B",
                ),
            )

    # --------- Time series sheets ----------
    def _build_timeseries_sheet(self, sheet_name, metric_key, units_label, number_format, change_format):
        metric = METRICS[metric_key]
        ws = self.wb.create_sheet(sheet_name)
        ws.sheet_view.showGridLines = False

        ws["A1"] = f"{sheet_name} — {metric} ({units_label})"
        self._font(ws["A1"], size=14, bold=True, color=COLOR_HEADER_BG)
        ws.row_dimensions[1].height = 22

        # Cap displayed dates to 100 most recent (full history still in Data_Raw).
        max_dates_shown = 100
        if len(self.dates) > max_dates_shown:
            display_idx_start = len(self.dates) - max_dates_shown
            ws["A2"] = (
                f"Showing the {max_dates_shown} most recent observations of {len(self.dates)} total. "
                f"Full history available on the Data_Raw tab."
            )
            ws.merge_cells("A2:H2")
            self._font(ws["A2"], color=COLOR_NOTE, italic=True)
        else:
            display_idx_start = 0

        display_dates = self.dates[display_idx_start:]
        header_row = 3
        date_start_col = 5
        for i, h in enumerate(["Region / Bucket", "Entity", "Latest", "YTD Δ"], start=1):
            self._hdr(ws.cell(row=header_row, column=i), h)
        for j, d in enumerate(display_dates):
            self._hdr(ws.cell(row=header_row, column=date_start_col + j), d.strftime("%Y-%m-%d"))

        ws.column_dimensions["A"].width = 14
        ws.column_dimensions["B"].width = 26
        ws.column_dimensions["C"].width = 11
        ws.column_dimensions["D"].width = 11
        for j in range(len(display_dates)):
            ws.column_dimensions[get_column_letter(date_start_col + j)].width = 11
        ws.freeze_panes = ws.cell(row=header_row + 1, column=date_start_col)

        # YTD is always anchored to the current real-world calendar year, not
        # to the latest data point. If you re-run in 2027 with stale 2026 data,
        # the denominator is end-of-2026; if you re-run in 2026, end-of-2025.
        latest_year = datetime.now().year
        ytd_idx = self.year_start_index(latest_year)
        # If ytd_idx is before the displayed window, drop YTD column references.
        ytd_in_window = (ytd_idx is not None) and (ytd_idx >= display_idx_start)
        ytd_ref_letter = (
            get_column_letter(date_start_col + (ytd_idx - display_idx_start))
            if ytd_in_window else None
        )
        latest_letter = (
            get_column_letter(date_start_col + len(display_dates) - 1)
            if display_dates else None
        )

        row = header_row + 1
        self._index_lbl(ws.cell(row=row, column=1), "Index")
        self._index_lbl(ws.cell(row=row, column=2), INDEX_NAME)
        self._write_series(ws, row, INDEX_NAME, metric, number_format, change_format,
                           date_start_col, display_idx_start, latest_letter, ytd_ref_letter)
        row += 2

        for region in REGION_ORDER:
            countries = self.countries_by_region.get(region, [])
            if not countries:
                continue
            agg = REGION_AGGREGATE[region]
            self._region_lbl(ws.cell(row=row, column=1), region)
            self._region_lbl(ws.cell(row=row, column=2), f"{agg} (rollup)")
            self._write_series(ws, row, agg, metric, number_format, change_format,
                               date_start_col, display_idx_start, latest_letter, ytd_ref_letter)
            row += 1
            for c in countries:
                ws.cell(row=row, column=1, value="").border = THIN_BORDER
                self._font(ws.cell(row=row, column=1))
                ws.cell(row=row, column=2, value=c).border = THIN_BORDER
                self._font(ws.cell(row=row, column=2))
                self._write_series(ws, row, c, metric, number_format, change_format,
                                   date_start_col, display_idx_start, latest_letter, ytd_ref_letter)
                row += 1
            row += 1

        ws.cell(row=row, column=1, value="By rating")
        self._font(ws.cell(row=row, column=1), bold=True, color=COLOR_HEADER_BG)
        row += 1
        for rating in self.ratings_present:
            self._rating_lbl(ws.cell(row=row, column=1), "Rating")
            self._rating_lbl(ws.cell(row=row, column=2), RATING_DISPLAY.get(rating, rating))
            self._write_series(ws, row, rating, metric, number_format, change_format,
                               date_start_col, display_idx_start, latest_letter, ytd_ref_letter)
            row += 1

        if display_dates:
            last = row - 1
            scale_col = "C"
            ws.conditional_formatting.add(
                f"{scale_col}{header_row + 1}:{scale_col}{last}",
                ColorScaleRule(
                    start_type="min", start_color="63BE7B",
                    mid_type="percentile", mid_value=50, mid_color="FFEB84",
                    end_type="max", end_color="F8696B",
                ),
            )

    def _write_series(self, ws, row, entity, metric, fmt, change_fmt,
                      date_start_col, display_idx_start, latest_letter, ytd_letter):
        vals = self.series.get((entity, metric), [None] * len(self.dates))
        for j, v in enumerate(vals[display_idx_start:]):
            cell = ws.cell(row=row, column=date_start_col + j)
            self._val(cell, v if v is not None else "", fmt=fmt, color=COLOR_HARDCODE)

        if latest_letter:
            c = ws.cell(row=row, column=3)
            c.value = f"={latest_letter}{row}"
            c.number_format = fmt
            self._font(c, bold=True)
            c.alignment = Alignment(horizontal="right")
            c.border = THIN_BORDER

        if latest_letter and ytd_letter:
            c = ws.cell(row=row, column=4)
            c.value = f'=IFERROR({latest_letter}{row}-{ytd_letter}{row}, "")'
            c.number_format = change_fmt
            self._font(c)
            c.alignment = Alignment(horizontal="right")
            c.border = THIN_BORDER

    # --------- TR_YTD ----------
    def build_tret_ytd(self):
        ws = self.wb.create_sheet("TR_YTD")
        ws.sheet_view.showGridLines = False
        ws["A1"] = f"Total Return — YTD {datetime.now().year} (% from end of {datetime.now().year - 1})"
        self._font(ws["A1"], size=14, bold=True, color=COLOR_HEADER_BG)
        ws.row_dimensions[1].height = 22

        max_dates = 100
        if len(self.dates) > max_dates:
            display_idx_start = len(self.dates) - max_dates
            ws["A2"] = f"Showing {max_dates} most recent of {len(self.dates)} obs."
            self._font(ws["A2"], color=COLOR_NOTE, italic=True)
        else:
            display_idx_start = 0
        display_dates = self.dates[display_idx_start:]

        header_row = 3
        date_start_col = 5
        for i, h in enumerate(["Region / Bucket", "Entity", "YTD Return", "Index Value"], start=1):
            self._hdr(ws.cell(row=header_row, column=i), h)
        for j, d in enumerate(display_dates):
            self._hdr(ws.cell(row=header_row, column=date_start_col + j), d.strftime("%Y-%m-%d"))

        ws.column_dimensions["A"].width = 14
        ws.column_dimensions["B"].width = 26
        ws.column_dimensions["C"].width = 12
        ws.column_dimensions["D"].width = 14
        for j in range(len(display_dates)):
            ws.column_dimensions[get_column_letter(date_start_col + j)].width = 11
        ws.freeze_panes = ws.cell(row=header_row + 1, column=date_start_col)

        # YTD is always anchored to the current real-world calendar year, not
        # to the latest data point. If you re-run in 2027 with stale 2026 data,
        # the denominator is end-of-2026; if you re-run in 2026, end-of-2025.
        latest_year = datetime.now().year
        ytd_idx = self.year_start_index(latest_year)
        ytd_in_window = (ytd_idx is not None) and (ytd_idx >= display_idx_start)
        ytd_letter = get_column_letter(date_start_col + (ytd_idx - display_idx_start)) if ytd_in_window else None
        latest_letter = get_column_letter(date_start_col + len(display_dates) - 1) if display_dates else None
        metric = METRICS["tret"]

        row = header_row + 1
        self._index_lbl(ws.cell(row=row, column=1), "Index")
        self._index_lbl(ws.cell(row=row, column=2), INDEX_NAME)
        self._write_tret(ws, row, INDEX_NAME, metric, date_start_col, display_idx_start, latest_letter, ytd_letter)
        row += 2

        for region in REGION_ORDER:
            countries = self.countries_by_region.get(region, [])
            if not countries:
                continue
            agg = REGION_AGGREGATE[region]
            self._region_lbl(ws.cell(row=row, column=1), region)
            self._region_lbl(ws.cell(row=row, column=2), f"{agg} (rollup)")
            self._write_tret(ws, row, agg, metric, date_start_col, display_idx_start, latest_letter, ytd_letter)
            row += 1
            for c in countries:
                ws.cell(row=row, column=1, value="").border = THIN_BORDER
                self._font(ws.cell(row=row, column=1))
                ws.cell(row=row, column=2, value=c).border = THIN_BORDER
                self._font(ws.cell(row=row, column=2))
                self._write_tret(ws, row, c, metric, date_start_col, display_idx_start, latest_letter, ytd_letter)
                row += 1
            row += 1

        ws.cell(row=row, column=1, value="By rating")
        self._font(ws.cell(row=row, column=1), bold=True, color=COLOR_HEADER_BG)
        row += 1
        for rating in self.ratings_present:
            self._rating_lbl(ws.cell(row=row, column=1), "Rating")
            self._rating_lbl(ws.cell(row=row, column=2), RATING_DISPLAY.get(rating, rating))
            self._write_tret(ws, row, rating, metric, date_start_col, display_idx_start, latest_letter, ytd_letter)
            row += 1

        if display_dates:
            last = row - 1
            ws.conditional_formatting.add(
                f"C{header_row + 1}:C{last}",
                ColorScaleRule(
                    start_type="min", start_color="F8696B",
                    mid_type="num", mid_value=0, mid_color="FFEB84",
                    end_type="max", end_color="63BE7B",
                ),
            )

    def _write_tret(self, ws, row, entity, metric, date_start_col, display_idx_start, latest_letter, ytd_letter):
        vals = self.series.get((entity, metric), [None] * len(self.dates))
        for j, v in enumerate(vals[display_idx_start:]):
            cell = ws.cell(row=row, column=date_start_col + j)
            self._val(cell, v if v is not None else "", fmt="0.00", color=COLOR_HARDCODE)

        if latest_letter and ytd_letter:
            c = ws.cell(row=row, column=3)
            c.value = f'=IFERROR(({latest_letter}{row}/{ytd_letter}{row})-1, "")'
            c.number_format = "0.00%;(0.00%);-"
            self._font(c, bold=True)
            c.alignment = Alignment(horizontal="right")
            c.border = THIN_BORDER

        if latest_letter:
            c = ws.cell(row=row, column=4)
            c.value = f"={latest_letter}{row}"
            c.number_format = "#,##0.00"
            self._font(c)
            c.alignment = Alignment(horizontal="right")
            c.border = THIN_BORDER

    # --------- By Rating ----------
    def build_by_rating(self):
        ws = self.wb.create_sheet("By_Rating")
        ws.sheet_view.showGridLines = False
        ws["A1"] = "EMBI Global — Composition by Credit Rating"
        self._font(ws["A1"], size=14, bold=True, color=COLOR_HEADER_BG)
        ws.row_dimensions[1].height = 22

        ws["A2"] = "Latest snapshot of each JPM rating bucket. YTD TR uses end-of-prior-year as denominator."
        ws.merge_cells("A2:F2")
        self._font(ws["A2"], color=COLOR_NOTE, italic=True)

        header_row = 4
        for i, h in enumerate(["Bucket", "Spread (bps)", "Yield (%)", "Total Return Idx", "YTD TR (%)", "Notes"], start=1):
            self._hdr(ws.cell(row=header_row, column=i), h)
        ws.column_dimensions["A"].width = 22
        ws.column_dimensions["B"].width = 14
        ws.column_dimensions["C"].width = 12
        ws.column_dimensions["D"].width = 18
        ws.column_dimensions["E"].width = 14
        ws.column_dimensions["F"].width = 30

        # YTD is always anchored to the current real-world calendar year, not
        # to the latest data point. If you re-run in 2027 with stale 2026 data,
        # the denominator is end-of-2026; if you re-run in 2026, end-of-2025.
        latest_year = datetime.now().year
        ytd_idx = self.year_start_index(latest_year)
        row = header_row + 1
        for rating in self.ratings_present:
            self._rating_lbl(ws.cell(row=row, column=1), RATING_DISPLAY.get(rating, rating))
            sp = self.latest_value(rating, METRICS["spread"])
            yd = self.latest_value(rating, METRICS["yield"])
            tr = self.latest_value(rating, METRICS["tret"])
            tr_start = self.value_at(rating, METRICS["tret"], ytd_idx) if ytd_idx is not None else None
            self._val(ws.cell(row=row, column=2), sp if sp is not None else "", fmt="0", color=COLOR_HARDCODE)
            self._val(ws.cell(row=row, column=3), yd if yd is not None else "", fmt="0.00", color=COLOR_HARDCODE)
            self._val(ws.cell(row=row, column=4), tr if tr is not None else "", fmt="#,##0.00", color=COLOR_HARDCODE)
            if tr is not None and tr_start not in (None, 0):
                self._val(ws.cell(row=row, column=5), (tr / tr_start) - 1, fmt="0.00%;(0.00%);-")
            else:
                self._val(ws.cell(row=row, column=5), "", fmt="0.00%;(0.00%);-")
            ws.cell(row=row, column=6, value="").border = THIN_BORDER
            row += 1
        ws.freeze_panes = "A5"

    # --------- Snapshot tab ----------
    def build_snapshot(self):
        ws = self.wb.create_sheet("Snapshot")
        ws.sheet_view.showGridLines = False
        ws["A1"] = "Per-country snapshot detail"
        self._font(ws["A1"], size=14, bold=True, color=COLOR_HEADER_BG)
        ws.row_dimensions[1].height = 22

        if self.snap_date:
            ws["A2"] = f"Snapshot date: {self.snap_date:%Y-%m-%d}"
        else:
            ws["A2"] = "No snapshot file loaded."
        self._font(ws["A2"], color=COLOR_NOTE, italic=True)

        # Pick the metrics that matter from the snapshot rows
        cols = [
            ("Region", 14, "region"),
            ("Country", 26, "country"),
            ("Mkt Cap %", 12, "Mkt Cap %"),
            ("S&P", 8, "Average S&P Rating"),
            ("Moody's", 9, "Average Moody Rating"),
            ("Fitch", 8, "Average Fitch Rating"),
            ("YTW", 10, "Yield to Worst"),
            ("STW (Trsy)", 11, "STW (Trsy)"),
            ("Z-Spread to Worst", 16, "Z Spread to Worst"),
            ("Spread Dur", 11, "Spread Duration"),
            ("Avg Life", 10, "Avr. Life"),
            ("MTD chg %", 11, "MTD Change (%)"),
            ("YTD chg %", 11, "YTD Change (%)"),
            ("# Issues", 10, "No. of Issues"),
            ("# Issuers", 10, "No. of Issuer"),
        ]
        header_row = 4
        for i, (h, w, _) in enumerate(cols, start=1):
            self._hdr(ws.cell(row=header_row, column=i), h)
            ws.column_dimensions[get_column_letter(i)].width = w

        row = header_row + 1
        for region in REGION_ORDER:
            countries = self.countries_by_region.get(region, [])
            if not countries:
                continue
            self._region_lbl(ws.cell(row=row, column=1), region)
            self._region_lbl(ws.cell(row=row, column=2), f"{region} subtotal")
            # Region rollup from snapshot
            snap = self.snap_data.get(f"REGION:{region}") or {}
            self._fill_snapshot_row(ws, row, snap, cols, label_style="region")
            row += 1
            for c in countries:
                snap = self.snap_data.get(c) or {}
                ws.cell(row=row, column=1, value="").border = THIN_BORDER
                self._font(ws.cell(row=row, column=1))
                ws.cell(row=row, column=2, value=c).border = THIN_BORDER
                self._font(ws.cell(row=row, column=2))
                self._fill_snapshot_row(ws, row, snap, cols, label_style="country")
                row += 1
            row += 1
        ws.freeze_panes = "C5"

    def _fill_snapshot_row(self, ws, row, snap, cols, *, label_style):
        for i, (h, _, key) in enumerate(cols, start=1):
            if key in ("region", "country"):
                continue  # already written by caller
            val = (snap.get(key) or "").strip() if snap else ""
            cell = ws.cell(row=row, column=i)
            cell.border = THIN_BORDER
            self._font(cell)
            cell.alignment = Alignment(horizontal="right" if key not in ("Average S&P Rating","Average Moody Rating","Average Fitch Rating") else "center")
            if val == "":
                cell.value = ""
                continue
            # Try numeric
            try:
                f = float(val)
                cell.value = f
                if key == "Mkt Cap %":
                    cell.number_format = "0.00"
                elif "Change" in key:
                    cell.number_format = "+0.00;-0.00;0.00"
                elif key in ("Yield to Worst", "STW (Trsy)", "Spread Duration", "Avr. Life"):
                    cell.number_format = "0.00"
                elif key == "Z Spread to Worst":
                    cell.number_format = "0"
                elif key in ("No. of Issues", "No. of Issuer"):
                    cell.number_format = "0"
                else:
                    cell.number_format = "0.00"
                self._font(cell, color=COLOR_HARDCODE)
            except ValueError:
                cell.value = val
                self._font(cell, color=COLOR_HARDCODE)

    # --------- Weights History ----------
    def build_weights_history(self):
        ws = self.wb.create_sheet("Weights_History")
        ws.sheet_view.showGridLines = False
        ws["A1"] = "Country weights — full history (monthly)"
        self._font(ws["A1"], size=14, bold=True, color=COLOR_HEADER_BG)
        ws.row_dimensions[1].height = 22

        if not self.weight_dates:
            ws["A2"] = "No weights data loaded."
            self._font(ws["A2"], color=COLOR_NOTE, italic=True)
            return

        # Cap displayed dates to last 60 (5 years monthly).
        max_dates = 60
        if len(self.weight_dates) > max_dates:
            ds = len(self.weight_dates) - max_dates
            ws["A2"] = f"Showing {max_dates} most recent of {len(self.weight_dates)} obs (full history below)."
        else:
            ds = 0
        self._font(ws["A2"], color=COLOR_NOTE, italic=True)
        display_dates = self.weight_dates[ds:]

        header_row = 4
        ws.column_dimensions["A"].width = 14
        ws.column_dimensions["B"].width = 26
        for j in range(len(display_dates)):
            ws.column_dimensions[get_column_letter(3 + j)].width = 9

        for i, h in enumerate(["Region", "Country"], start=1):
            self._hdr(ws.cell(row=header_row, column=i), h)
        for j, d in enumerate(display_dates):
            self._hdr(ws.cell(row=header_row, column=3 + j), d.strftime("%Y-%m"))

        ws.freeze_panes = ws.cell(row=header_row + 1, column=3)
        row = header_row + 1
        for region in REGION_ORDER:
            countries = self.countries_by_region.get(region, [])
            if not countries:
                continue
            for c in countries:
                self._region_lbl(ws.cell(row=row, column=1), region) if c == countries[0] else (
                    ws.cell(row=row, column=1, value="").__setattr__("border", THIN_BORDER)
                )
                ws.cell(row=row, column=2, value=c).border = THIN_BORDER
                self._font(ws.cell(row=row, column=2))
                vals = self.weight_series.get(c, [None] * len(self.weight_dates))
                for j, v in enumerate(vals[ds:]):
                    cell = ws.cell(row=row, column=3 + j)
                    if v is None:
                        cell.value = ""
                    else:
                        cell.value = v
                        cell.number_format = "0.00"
                        self._font(cell, color=COLOR_HARDCODE)
                    cell.border = THIN_BORDER
                row += 1

    # --------- LatAm Focus ----------
    def build_latam_focus(self):
        ws = self.wb.create_sheet("LatAm_Focus")
        ws.sheet_view.showGridLines = False
        focus_present = [c for c in LATAM_FOCUS if (c in {e for (e, _) in self.series.keys()}) or (c in self.snap_data)]
        latam_others = [c for c in self.countries_by_region["LatAm"] if c not in LATAM_FOCUS]
        peer_aggregates = [
            ("LatAm rollup",   REGION_AGGREGATE["LatAm"]),
            ("Asia rollup",    REGION_AGGREGATE["Asia"]),
            ("Europe rollup",  REGION_AGGREGATE["Europe"]),
            ("Africa rollup",  REGION_AGGREGATE["Africa"]),
            ("GCC rollup",     REGION_AGGREGATE["GCC"]),
            ("EMBI Global",    INDEX_NAME),
        ]

        ws["A1"] = "LatAm Focus — Performance & Spread Snapshot"
        self._font(ws["A1"], size=14, bold=True, color=COLOR_HEADER_BG)
        ws.row_dimensions[1].height = 22
        ws["A2"] = ("Top: focus credits. Middle: peer rollups. Bottom: other LatAm constituents. "
                    "All YTD figures use end-of-prior-year as denominator.")
        ws.merge_cells("A2:H2")
        self._font(ws["A2"], color=COLOR_NOTE, italic=True)
        ws["A2"].alignment = Alignment(wrap_text=True, vertical="top")
        ws.row_dimensions[2].height = 32

        header_row = 4
        cols = [("Bucket", 14), ("Entity", 22), ("Spread (bps)", 12), ("YTD Δ Spread", 12),
                ("Yield (%)", 11), ("YTD Δ Yield", 12), ("TR Index", 12), ("YTD TR (%)", 12)]
        for i, (h, w) in enumerate(cols, start=1):
            self._hdr(ws.cell(row=header_row, column=i), h)
            ws.column_dimensions[get_column_letter(i)].width = w

        # YTD is always anchored to the current real-world calendar year, not
        # to the latest data point. If you re-run in 2027 with stale 2026 data,
        # the denominator is end-of-2026; if you re-run in 2026, end-of-2025.
        latest_year = datetime.now().year
        ytd_idx = self.year_start_index(latest_year)

        def _snap(row, bucket, entity, *, label_style):
            if label_style == "focus":
                self._region_lbl(ws.cell(row=row, column=1), bucket)
                cell = ws.cell(row=row, column=2, value=entity)
                self._font(cell, bold=True)
                cell.fill = PatternFill("solid", start_color=COLOR_REGION_BG)
                cell.border = THIN_BORDER
            elif label_style == "peer":
                self._index_lbl(ws.cell(row=row, column=1), bucket)
                self._index_lbl(ws.cell(row=row, column=2), entity)
            else:
                ws.cell(row=row, column=1, value=bucket).border = THIN_BORDER
                self._font(ws.cell(row=row, column=1))
                ws.cell(row=row, column=2, value=entity).border = THIN_BORDER
                self._font(ws.cell(row=row, column=2))

            sp_now = self.latest_value(entity, METRICS["spread"])
            sp_start = self.value_at(entity, METRICS["spread"], ytd_idx) if ytd_idx is not None else None
            yd_now = self.latest_value(entity, METRICS["yield"])
            yd_start = self.value_at(entity, METRICS["yield"], ytd_idx) if ytd_idx is not None else None
            tr_now = self.latest_value(entity, METRICS["tret"])
            tr_start = self.value_at(entity, METRICS["tret"], ytd_idx) if ytd_idx is not None else None
            self._val(ws.cell(row=row, column=3), sp_now if sp_now is not None else "", fmt="0", color=COLOR_HARDCODE)
            self._val(ws.cell(row=row, column=4), (sp_now - sp_start) if sp_now is not None and sp_start is not None else "", fmt="+0;-0;0")
            self._val(ws.cell(row=row, column=5), yd_now if yd_now is not None else "", fmt="0.00", color=COLOR_HARDCODE)
            self._val(ws.cell(row=row, column=6), (yd_now - yd_start) if yd_now is not None and yd_start is not None else "", fmt="+0.00;-0.00;0.00")
            self._val(ws.cell(row=row, column=7), tr_now if tr_now is not None else "", fmt="#,##0.00", color=COLOR_HARDCODE)
            if tr_now is not None and tr_start not in (None, 0):
                self._val(ws.cell(row=row, column=8), (tr_now / tr_start) - 1, fmt="0.00%;(0.00%);-")
            else:
                self._val(ws.cell(row=row, column=8), "", fmt="0.00%;(0.00%);-")

        row = header_row + 1
        ws.cell(row=row, column=1, value="FOCUS CREDITS")
        self._font(ws.cell(row=row, column=1), bold=True, color=COLOR_HEADER_BG)
        row += 1
        for c in focus_present:
            _snap(row, "LatAm focus", c, label_style="focus")
            row += 1
        row += 1
        ws.cell(row=row, column=1, value="PEER BENCHMARKS")
        self._font(ws.cell(row=row, column=1), bold=True, color=COLOR_HEADER_BG)
        row += 1
        for label, ent in peer_aggregates:
            _snap(row, label, ent, label_style="peer")
            row += 1
        row += 1
        if latam_others:
            ws.cell(row=row, column=1, value="OTHER LATAM CONSTITUENTS")
            self._font(ws.cell(row=row, column=1), bold=True, color=COLOR_HEADER_BG)
            row += 1
            for c in latam_others:
                _snap(row, "LatAm", c, label_style="other")
                row += 1

        snap_last_row = row - 1
        ws.conditional_formatting.add(
            f"D{header_row + 1}:D{snap_last_row}",
            ColorScaleRule(start_type="min", start_color="63BE7B",
                           mid_type="num", mid_value=0, mid_color="FFEB84",
                           end_type="max", end_color="F8696B"),
        )
        ws.conditional_formatting.add(
            f"H{header_row + 1}:H{snap_last_row}",
            ColorScaleRule(start_type="min", start_color="F8696B",
                           mid_type="num", mid_value=0, mid_color="FFEB84",
                           end_type="max", end_color="63BE7B"),
        )
        ws.freeze_panes = ws.cell(row=header_row + 1, column=3)

        # ---- chart blocks ----
        chart_block = snap_last_row + 5
        ws.cell(row=chart_block, column=1, value="Chart data — focus + benchmarks")
        self._font(ws.cell(row=chart_block, column=1), bold=True, color=COLOR_NOTE, italic=True)

        # Limit chart data to last 60 obs to keep chart readable
        chart_dates = self.dates[-60:] if len(self.dates) > 60 else self.dates[:]
        offset_idx = len(self.dates) - len(chart_dates)
        chart_columns = focus_present + [REGION_AGGREGATE["LatAm"], INDEX_NAME]
        chart_labels  = focus_present + ["LatAm Region", "EMBI Global"]

        def _block(start_row, title, metric, fmt, rebase):
            ws.cell(row=start_row, column=1, value=title)
            self._font(ws.cell(row=start_row, column=1), bold=True)
            ws.cell(row=start_row + 1, column=1, value="Date")
            self._font(ws.cell(row=start_row + 1, column=1), bold=True)
            for j, lbl in enumerate(chart_labels):
                ws.cell(row=start_row + 1, column=2 + j, value=lbl)
                self._font(ws.cell(row=start_row + 1, column=2 + j), bold=True)
            bases: Dict[str, Optional[float]] = {}
            if rebase:
                for ent in chart_columns:
                    for v in self.series.get((ent, metric), [])[offset_idx:]:
                        if v is not None:
                            bases[ent] = v
                            break
                    else:
                        bases[ent] = None
            for i, d in enumerate(chart_dates):
                ws.cell(row=start_row + 2 + i, column=1, value=d).number_format = "yyyy-mm-dd"
                for j, ent in enumerate(chart_columns):
                    v = self.value_at(ent, metric, offset_idx + i)
                    if v is None:
                        continue
                    if rebase:
                        base = bases.get(ent)
                        if not base:
                            continue
                        v = (v / base) * 100
                    cell = ws.cell(row=start_row + 2 + i, column=2 + j, value=v)
                    cell.number_format = fmt
            return start_row + 1, start_row + 1 + len(chart_dates), start_row + 2 + len(chart_dates)

        spr_h, spr_l, nxt = _block(chart_block + 2, "Z-Spread (bps)", METRICS["spread"], "0", False)
        yld_h, yld_l, nxt = _block(nxt + 3, "Yield to Maturity (%)", METRICS["yield"], "0.00", False)
        tr_h, tr_l, _ = _block(nxt + 3, "Cumulative Total Return — rebased to 100", METRICS["tret"], "0.00", True)

        def _line(title, ytitle, hdr, last):
            ch = LineChart()
            ch.title = title
            ch.style = 12
            ch.y_axis.title = ytitle
            ch.height = 11
            ch.width = 24
            data = Reference(ws, min_col=2, max_col=1 + len(chart_columns), min_row=hdr, max_row=last)
            cats = Reference(ws, min_col=1, min_row=hdr + 1, max_row=last)
            ch.add_data(data, titles_from_data=True)
            ch.set_categories(cats)
            return ch

        ws.add_chart(_line("Z-Spread — focus vs LatAm vs EMBI Global", "bps", spr_h, spr_l), f"J{header_row}")
        ws.add_chart(_line("Yield to Maturity — focus vs benchmarks", "%", yld_h, yld_l), f"J{header_row + 22}")
        ws.add_chart(_line("Cumulative Total Return (rebased = 100)", "Index", tr_h, tr_l), f"J{header_row + 44}")

    # --------- Forecast Simulator ----------
    def build_forecast(self):
        """Interactive forecast monitor.

        Model:
          ΔSpread_drift = β_UST · ΔUST + β_Oil · ΔOil + β_Rating · ΔRating
                          + (μ_monthly · horizon_months)
          σ_spread_h    = σ_monthly · √horizon_months · vol_multiplier
          Spread P50    = current_spread + ΔSpread_drift
          Spread Pα     = P50 + Φ⁻¹(α) · σ_spread_h        (analytical normal)
          ΔYield_drift  = ΔUST/100 + ΔSpread_drift/100
          Yield P50     = current_yield + ΔYield_drift
          TR P50        = (current_yield · h/12) − duration · ΔYield_drift

        Rationale: the analytical-normal percentile approach is mathematically
        equivalent to a Monte Carlo simulation with infinite paths under
        normality. It updates instantly, has no F9-flicker, and is therefore
        the right choice for a "monitor" use case where the user wants to
        plug in assumptions and read the answer.
        """
        ws = self.wb.create_sheet("Forecast")
        ws.sheet_view.showGridLines = False

        # ----- 1. Compute historical statistics for EMBI Global -----
        spread_series = self.series.get((INDEX_NAME, METRICS["spread"])) or []
        yield_series  = self.series.get((INDEX_NAME, METRICS["yield"])) or []

        def _diffs(vals):
            return [vals[i] - vals[i - 1]
                    for i in range(1, len(vals))
                    if vals[i] is not None and vals[i - 1] is not None]

        spread_diffs = _diffs(spread_series)
        yield_diffs  = _diffs(yield_series)
        sp_mean = statistics.mean(spread_diffs) if spread_diffs else 0.0
        sp_std  = statistics.stdev(spread_diffs) if len(spread_diffs) > 1 else 0.0
        yd_mean = statistics.mean(yield_diffs)  if yield_diffs else 0.0
        yd_std  = statistics.stdev(yield_diffs) if len(yield_diffs) > 1 else 0.0

        cur_spread = self.latest_value(INDEX_NAME, METRICS["spread"]) or 0.0
        cur_yield  = self.latest_value(INDEX_NAME, METRICS["yield"]) or 0.0

        # Spread duration — pull from snapshot if available, otherwise a
        # reasonable EMBI Global default.
        snap_idx = self.snap_data.get("INDEX:EMBIGD") or {}
        try:
            duration = float(snap_idx.get("Spread Duration") or "")
        except (TypeError, ValueError):
            duration = 5.7  # historical EMBI Global avg

        # Default macro sensitivities — practitioner-standard, editable by user.
        # Oil was deliberately removed: EMBI Global mixes oil exporters (Saudi,
        # UAE, Mexico, Colombia, Nigeria, Angola) with major importers (China,
        # India, Turkey, Indonesia). The aggregate sensitivity is small, noisy,
        # and policy-regime-dependent. DXY and VIX are far more robust.
        beta_ust    = -0.15  # bps spread per +1bp UST (IG-heavy index, mild neg corr)
        beta_dxy    =  6.0   # bps spread per +1% DXY rally (USD strength → EM stress)
        beta_vix    =  4.0   # bps spread per +1 VIX point (risk-off → wider spreads)
        beta_rating = -50.0  # bps spread per +1 notch rating improvement

        # ----- 2. Layout -----
        ws.column_dimensions["A"].width = 4
        ws.column_dimensions["B"].width = 42
        ws.column_dimensions["C"].width = 14
        ws.column_dimensions["D"].width = 14
        ws.column_dimensions["E"].width = 14
        ws.column_dimensions["F"].width = 4
        ws.column_dimensions["G"].width = 38

        ws["B1"] = "Forecast Simulator — EMBI Global"
        self._font(ws["B1"], size=18, bold=True, color=COLOR_HEADER_BG)
        ws.row_dimensions[1].height = 26
        ws["B2"] = ("Edit the YELLOW cells. Everything else updates automatically. "
                    "Forecast = analytical Monte Carlo (normal-distribution percentile "
                    "bands) calibrated on the historical EMBI Global Z-Spread vol.")
        ws.merge_cells("B2:G2")
        self._font(ws["B2"], italic=True, color=COLOR_NOTE)
        ws["B2"].alignment = Alignment(wrap_text=True, vertical="top")
        ws.row_dimensions[2].height = 30

        def _section_header(row, text):
            cell = ws.cell(row=row, column=2, value=text)
            self._font(cell, bold=True, size=12, color=COLOR_HEADER_BG)
            cell.fill = PatternFill("solid", start_color=COLOR_REGION_BG)
            ws.cell(row=row, column=3).fill = PatternFill("solid", start_color=COLOR_REGION_BG)
            ws.cell(row=row, column=4).fill = PatternFill("solid", start_color=COLOR_REGION_BG)
            ws.cell(row=row, column=5).fill = PatternFill("solid", start_color=COLOR_REGION_BG)

        def _input_cell(row, label, value, fmt="0", note=""):
            ws.cell(row=row, column=2, value=label)
            self._font(ws.cell(row=row, column=2))
            cell = ws.cell(row=row, column=3, value=value)
            cell.fill = PatternFill("solid", start_color=COLOR_INPUT_BG)
            cell.number_format = fmt
            self._font(cell, color=COLOR_HARDCODE, bold=True)
            cell.alignment = Alignment(horizontal="right")
            cell.border = THIN_BORDER
            if note:
                ws.cell(row=row, column=7, value=note)
                self._font(ws.cell(row=row, column=7), italic=True, color=COLOR_NOTE, size=9)
                ws.cell(row=row, column=7).alignment = Alignment(wrap_text=True, vertical="center")
            return cell

        def _const_cell(row, label, value, fmt="0.00"):
            ws.cell(row=row, column=2, value=label)
            self._font(ws.cell(row=row, column=2))
            cell = ws.cell(row=row, column=3, value=value)
            cell.number_format = fmt
            self._font(cell, color=COLOR_HARDCODE)
            cell.alignment = Alignment(horizontal="right")
            cell.border = THIN_BORDER
            return cell

        # ----- 3. Assumption inputs (yellow) -----
        _section_header(4, "YOUR ASSUMPTIONS  (yellow cells — edit these)")
        ust_input    = _input_cell(5, "UST 10Y change over the horizon (bps)", 0,
                                   fmt="+0;-0;0",
                                   note="Positive = rates rise. e.g. +50 = 50bp UST sell-off.")
        dxy_input    = _input_cell(6, "DXY (US Dollar Index) change (%)", 0,
                                   fmt="+0.0%;-0.0%;0.0%",
                                   note="Positive = USD strength. Format as decimal (0.05 = +5%). USD strength tightens dollar funding → wider EM spreads.")
        vix_input    = _input_cell(7, "VIX change (points)", 0,
                                   fmt="+0;-0;0",
                                   note="Absolute change in VIX. e.g. +10 = VIX moves from 18 to 28 (risk-off shock).")
        rating_input = _input_cell(8, "EM rating drift (S&P notches; +1 = upgrade)", 0,
                                   fmt="+0;-0;0",
                                   note="Average index-wide rating drift. +1 ≈ index moves up one notch.")
        vol_input    = _input_cell(9, "Volatility multiplier (1.0 = baseline; 2.0 = stress)", 1.0,
                                   fmt="0.0",
                                   note="Variance scaling — independent of the mean drivers above. Use >1 for fatter tails / regime-shift scenarios.")

        # ----- 4. Current state (read-only) -----
        _section_header(11, "CURRENT STATE  (latest data — read only)")
        ws.cell(row=12, column=2, value="Latest data date")
        self._font(ws.cell(row=12, column=2))
        ws.cell(row=12, column=3, value=self.dates[-1] if self.dates else "")
        ws.cell(row=12, column=3).number_format = "yyyy-mm-dd"
        self._font(ws.cell(row=12, column=3), color=COLOR_HARDCODE)
        ws.cell(row=12, column=3).alignment = Alignment(horizontal="right")
        ws.cell(row=12, column=3).border = THIN_BORDER

        cur_spread_cell = _const_cell(13, "EMBI Global Z-Spread (bps)", cur_spread, fmt="0")
        cur_yield_cell  = _const_cell(14, "EMBI Global Yield to Maturity (%)", cur_yield / 100, fmt="0.00%")
        # Implied UST = yield - spread/10000 (since spread in bps and yield in %)
        ws.cell(row=15, column=2, value="Implied UST 10Y (%)")
        self._font(ws.cell(row=15, column=2))
        ws.cell(row=15, column=3, value=f"=C14-C13/10000")
        ws.cell(row=15, column=3).number_format = "0.00%"
        self._font(ws.cell(row=15, column=3))
        ws.cell(row=15, column=3).alignment = Alignment(horizontal="right")
        ws.cell(row=15, column=3).border = THIN_BORDER

        dur_cell    = _const_cell(16, "Spread duration (years)", duration, fmt="0.00")
        sp_mean_cell= _const_cell(17, "Historical Δspread monthly mean (bps)", sp_mean, fmt="+0.0;-0.0;0.0")
        sp_std_cell = _const_cell(18, "Historical Δspread monthly stdev (bps)", sp_std, fmt="0.0")
        yd_std_cell = _const_cell(19, "Historical Δyield monthly stdev (%)", yd_std / 100, fmt="0.00%")
        ws.cell(row=19, column=7, value=f"Sample: {len(spread_diffs)} monthly observations of EMBI Global.")
        self._font(ws.cell(row=19, column=7), italic=True, color=COLOR_NOTE, size=9)

        # ----- 5. Sensitivities (advanced; editable) -----
        _section_header(21, "MACRO SENSITIVITIES  (defaults from practitioner research — override if you have a strong view)")
        beta_ust_cell = _const_cell(22, "β: bps spread move per +1bp UST",          beta_ust,    fmt="0.000")
        beta_dxy_cell = _const_cell(23, "β: bps spread move per +1% DXY rally",      beta_dxy,    fmt="0.0")
        beta_vix_cell = _const_cell(24, "β: bps spread move per +1 VIX point",       beta_vix,    fmt="0.0")
        beta_rat_cell = _const_cell(25, "β: bps spread move per +1 rating notch",    beta_rating, fmt="0")
        ws.cell(row=22, column=7, value="EM IG-heavy index: rising UST tends to compress spread modestly. Empirical band: -0.05 to -0.30.")
        ws.cell(row=23, column=7, value="Best single FX driver of EM credit. USD strength tightens dollar funding → wider spreads. Empirical band: +3 to +9 bps/%.")
        ws.cell(row=24, column=7, value="Pure risk-off proxy. Captures sentiment regime independently of rates/FX. Empirical band: +2 to +7 bps/point.")
        ws.cell(row=25, column=7, value="A 1-notch broad rating drift moves spreads ~50bps. Empirical band: 30 to 80.")
        for r in (22, 23, 24, 25):
            self._font(ws.cell(row=r, column=7), italic=True, color=COLOR_NOTE, size=9)
            ws.cell(row=r, column=7).alignment = Alignment(wrap_text=True, vertical="center")

        # ----- 6. Forecast output -----
        _section_header(27, "FORECAST  (3 / 6 / 12 months — auto-calculated)")
        # Header row
        ws.cell(row=28, column=2, value="Metric")
        for i, h in enumerate(["3 months", "6 months", "12 months"]):
            self._hdr(ws.cell(row=28, column=3 + i), h)
        self._font(ws.cell(row=28, column=2), bold=True)
        ws.cell(row=28, column=2).fill = PatternFill("solid", start_color=COLOR_HEADER_BG)
        ws.cell(row=28, column=2).font = Font(name=FONT_NAME, size=10, bold=True, color=COLOR_HEADER_FG)

        horizons = [3, 6, 12]
        # Z-score for the 5%/95% percentile
        Z95 = 1.6449
        Z75 = 0.6745

        # Helper to write a row of forecast formulas across the 3 horizon cols
        def _row(row, label, fmt, formula_factory, *, bold=False):
            ws.cell(row=row, column=2, value=label)
            self._font(ws.cell(row=row, column=2), bold=bold)
            for i, h in enumerate(horizons):
                cell = ws.cell(row=row, column=3 + i, value=formula_factory(h))
                cell.number_format = fmt
                self._font(cell, bold=bold)
                cell.alignment = Alignment(horizontal="right")
                cell.border = THIN_BORDER
                if bold:
                    cell.fill = PatternFill("solid", start_color=COLOR_INDEX_BG)

        # Cell references shorthand. Inputs/state/sensitivities now occupy:
        #   C5  UST input (bps)        |  C13 current spread (bps)
        #   C6  DXY input (decimal %)  |  C14 current yield (decimal)
        #   C7  VIX input (points)     |  C16 spread duration
        #   C8  Rating input (notches) |  C17 historical mean Δspread (bps)
        #   C9  Volatility multiplier  |  C18 historical std Δspread (bps)
        #   C22 β_UST | C23 β_DXY | C24 β_VIX | C25 β_Rating
        UST_C, DXY_C, VIX_C, RAT_C, VOL_C = "$C$5", "$C$6", "$C$7", "$C$8", "$C$9"
        SPR_C, YLD_C = "$C$13", "$C$14"
        DUR_C = "$C$16"
        SPMEAN_C, SPSTD_C = "$C$17", "$C$18"
        BUST_C, BDXY_C, BVIX_C, BRAT_C = "$C$22", "$C$23", "$C$24", "$C$25"

        # ΔSpread drift formula (in bps, over horizon h):
        #   = β_UST    × ΔUST_bps
        #   + β_DXY    × ΔDXY_pct   (DXY input is decimal; ×100 → percent points)
        #   + β_VIX    × ΔVIX_points
        #   + β_Rating × ΔRating_notches
        #   + μ_monthly × horizon_months
        def _drift(h):
            return (f"={BUST_C}*{UST_C}"
                    f"+{BDXY_C}*{DXY_C}*100"
                    f"+{BVIX_C}*{VIX_C}"
                    f"+{BRAT_C}*{RAT_C}"
                    f"+{SPMEAN_C}*{h}")

        # σ spread for horizon h
        def _sigma(h):
            return f"={SPSTD_C}*SQRT({h})*{VOL_C}"

        # Spread P50/P5/P95
        def _sp_p50(h): return f"={SPR_C}+({_drift(h)[1:]})"
        def _sp_p5(h):  return f"={SPR_C}+({_drift(h)[1:]})-{Z95}*({_sigma(h)[1:]})"
        def _sp_p95(h): return f"={SPR_C}+({_drift(h)[1:]})+{Z95}*({_sigma(h)[1:]})"
        def _sp_p25(h): return f"={SPR_C}+({_drift(h)[1:]})-{Z75}*({_sigma(h)[1:]})"
        def _sp_p75(h): return f"={SPR_C}+({_drift(h)[1:]})+{Z75}*({_sigma(h)[1:]})"

        # Yield P50/P5/P95: yield_now + ΔUST/10000 (bps→pct decimal) + ΔSpread_drift/10000
        def _yd_p50(h):
            return (f"={YLD_C}+{UST_C}/10000+(({_drift(h)[1:]})/10000)")
        def _yd_p5(h):
            return (f"={YLD_C}+{UST_C}/10000+(({_drift(h)[1:]})/10000)"
                    f"-{Z95}*({_sigma(h)[1:]})/10000")
        def _yd_p95(h):
            return (f"={YLD_C}+{UST_C}/10000+(({_drift(h)[1:]})/10000)"
                    f"+{Z95}*({_sigma(h)[1:]})/10000")

        # TR (decimal return) ≈ carry − duration × ΔYield_in_decimal
        # carry  = current_yield_decimal × h/12
        # ΔYield_in_decimal = ΔUST_bps/10000 + Δspread_drift_bps/10000
        # σ_TR (decimal)    = duration × σ_spread_h / 10000
        def _tr_p50(h):
            return (f"={YLD_C}*{h}/12-{DUR_C}*({UST_C}/10000+(({_drift(h)[1:]})/10000))")
        def _tr_p5(h):
            return (f"={YLD_C}*{h}/12-{DUR_C}*({UST_C}/10000+(({_drift(h)[1:]})/10000))"
                    f"-{Z95}*{DUR_C}*({_sigma(h)[1:]})/10000")
        def _tr_p95(h):
            return (f"={YLD_C}*{h}/12-{DUR_C}*({UST_C}/10000+(({_drift(h)[1:]})/10000))"
                    f"+{Z95}*{DUR_C}*({_sigma(h)[1:]})/10000")

        _row(29, "Median spread (bps)",       "0",            _sp_p50, bold=True)
        _row(30, "5th percentile spread (bps)",  "0",         _sp_p5)
        _row(31, "95th percentile spread (bps)", "0",         _sp_p95)
        _row(32, "Median yield (%)",          "0.00%",        _yd_p50, bold=True)
        _row(33, "5th percentile yield (%)",  "0.00%",        _yd_p5)
        _row(34, "95th percentile yield (%)", "0.00%",        _yd_p95)
        _row(35, "Median total return (%)",   "+0.00%;-0.00%;0.00%", _tr_p50, bold=True)
        _row(36, "5th percentile TR (%)",     "+0.00%;-0.00%;0.00%", _tr_p5)
        _row(37, "95th percentile TR (%)",    "+0.00%;-0.00%;0.00%", _tr_p95)

        # ----- 7. Stress matrix: TR median across UST shocks -----
        _section_header(40, "STRESS TABLE — Median TR forecast across UST scenarios")
        ws.cell(row=41, column=2, value="UST shock (bps)")
        self._font(ws.cell(row=41, column=2), bold=True)
        for i, h in enumerate(["3 months", "6 months", "12 months"]):
            self._hdr(ws.cell(row=41, column=3 + i), h)

        ust_scenarios = [-100, -50, 0, +50, +100]
        for r_offset, ust_shk in enumerate(ust_scenarios):
            r = 42 + r_offset
            label = f"{ust_shk:+d} bps" if ust_shk != 0 else "0 bps (UST flat)"
            ws.cell(row=r, column=2, value=label)
            self._font(ws.cell(row=r, column=2))
            ws.cell(row=r, column=2).border = THIN_BORDER
            for i, h in enumerate(horizons):
                # TR at given UST shock = yield·h/12 - dur·(ΔUST + Δspread)/10000
                # Δspread holds DXY/VIX/rating at user-specified levels.
                drift_at_shk = (f"({BUST_C}*({ust_shk})"
                                f"+{BDXY_C}*{DXY_C}*100"
                                f"+{BVIX_C}*{VIX_C}"
                                f"+{BRAT_C}*{RAT_C}"
                                f"+{SPMEAN_C}*{h})")
                formula = (f"={YLD_C}*{h}/12-{DUR_C}*(({ust_shk})/10000+({drift_at_shk}/10000))")
                cell = ws.cell(row=r, column=3 + i, value=formula)
                cell.number_format = "+0.00%;-0.00%;0.00%"
                self._font(cell)
                cell.alignment = Alignment(horizontal="right")
                cell.border = THIN_BORDER

        # Color scale on stress matrix
        ws.conditional_formatting.add(
            "C42:E46",
            ColorScaleRule(
                start_type="min", start_color="F8696B",
                mid_type="num", mid_value=0, mid_color="FFEB84",
                end_type="max", end_color="63BE7B",
            ),
        )

        # ----- 8. Fan chart data + chart -----
        # Build a 13-month projection (month 0 to 12) for each percentile band.
        # Place data far to the right so the user doesn't see the helper block.
        chart_data_col = 11  # column K
        ws.cell(row=4, column=chart_data_col, value="Forecast trajectory (helper data — feeds chart)")
        self._font(ws.cell(row=4, column=chart_data_col), italic=True, color=COLOR_NOTE, size=9)
        ws.cell(row=5, column=chart_data_col, value="Month")
        for j, lbl in enumerate(["P5", "P25", "Median", "P75", "P95"]):
            ws.cell(row=5, column=chart_data_col + 1 + j, value=lbl)
            self._font(ws.cell(row=5, column=chart_data_col + 1 + j), bold=True)
        for m in range(0, 13):
            r = 6 + m
            ws.cell(row=r, column=chart_data_col, value=m)
            ws.cell(row=r, column=chart_data_col).number_format = "0"
            # Use formulas referencing the input cells — chart updates with assumptions.
            for j, z_factor, sign in [(0, Z95, -1), (1, Z75, -1), (2, 0, 0), (3, Z75, 1), (4, Z95, 1)]:
                drift = (f"({BUST_C}*{UST_C}"
                         f"+{BDXY_C}*{DXY_C}*100"
                         f"+{BVIX_C}*{VIX_C}"
                         f"+{BRAT_C}*{RAT_C}"
                         f"+{SPMEAN_C}*{m})")
                sigma = f"({SPSTD_C}*SQRT({m})*{VOL_C})" if m > 0 else "0"
                if z_factor == 0:
                    formula = f"={SPR_C}+{drift}"
                else:
                    formula = f"={SPR_C}+{drift}{'+' if sign > 0 else '-'}{z_factor}*{sigma}"
                cell = ws.cell(row=r, column=chart_data_col + 1 + j, value=formula)
                cell.number_format = "0"
                self._font(cell, color=COLOR_NOTE, size=9)

        # Build the line chart (5 lines for the percentile bands)
        fan = LineChart()
        fan.title = "EMBI Global Z-Spread fan — projected (bps)"
        fan.style = 12
        fan.y_axis.title = "bps"
        fan.x_axis.title = "Months ahead"
        fan.height = 11
        fan.width = 22
        data_ref = Reference(
            ws, min_col=chart_data_col + 1, max_col=chart_data_col + 5,
            min_row=5, max_row=5 + 13,
        )
        cats_ref = Reference(
            ws, min_col=chart_data_col, min_row=6, max_row=5 + 13,
        )
        fan.add_data(data_ref, titles_from_data=True)
        fan.set_categories(cats_ref)
        ws.add_chart(fan, "B49")

        # ----- 9. Methodology / disclaimer footer -----
        ws.cell(row=72, column=2, value="Methodology")
        self._font(ws.cell(row=72, column=2), bold=True, size=11, color=COLOR_HEADER_BG)
        notes = [
            ("Driver selection",
             "UST, DXY, VIX, and rating drift are the four variables with the strongest, most "
             "documented impact on aggregate EM credit spreads. Oil was deliberately excluded — "
             "the EMBI Global universe contains both major oil exporters (Saudi, UAE, Mexico, "
             "Nigeria, Angola) and major importers (China, India, Turkey, Indonesia), so oil's "
             "aggregate effect is small, noisy, and policy-dependent. DXY is a far cleaner FX "
             "driver because USD strength tightens dollar-funding for the entire EM dollar-debt "
             "complex regardless of country-level commodity exposure."),
            ("Vol scaling",
             "Spread vol scales with √horizon (Brownian assumption)."),
            ("Distribution",
             "Returns assumed normal. Real spread distributions are fatter-tailed; treat the 5/95 bands "
             "as approximate. For a true tail-risk view, set Volatility multiplier to 1.5–2.0."),
            ("Calibration",
             "Mean and σ are estimated from the EMBI Global Z-Spread monthly time series in the loaded "
             "returns file. As you re-run the script with more data, the calibration sample grows and "
             "the model becomes more reliable."),
            ("Sensitivity coefficients",
             "Default β values are practitioner-standard, not estimated in-sample. β_DXY ≈ +6 bps/% "
             "and β_VIX ≈ +4 bps/point are well within the empirical ranges reported in IMF / BIS push-"
             "pull literature and Fed research notes. If you have your own regression, override rows 22–25."),
            ("What this is NOT",
             "Not a trading model. Not a Bloomberg-grade analytic. A scenario monitor intended to "
             "translate a macro view into a directional EMBIG forecast."),
        ]
        for i, (k, v) in enumerate(notes, start=73):
            ws.cell(row=i, column=2, value=k)
            self._font(ws.cell(row=i, column=2), bold=True)
            ws.cell(row=i, column=2).alignment = Alignment(vertical="top")
            ws.cell(row=i, column=3, value=v)
            self._font(ws.cell(row=i, column=3))
            ws.cell(row=i, column=3).alignment = Alignment(wrap_text=True, vertical="top")
            ws.merge_cells(start_row=i, start_column=3, end_row=i, end_column=7)
            ws.row_dimensions[i].height = max(18, 14 * (1 + len(v) // 80))

    # --------- Charts (index-wide) ----------
    def build_charts(self):
        ws = self.wb.create_sheet("Charts")
        ws.sheet_view.showGridLines = False
        ws["A1"] = "Charts — refreshed automatically each script run"
        self._font(ws["A1"], size=14, bold=True, color=COLOR_HEADER_BG)
        ws.row_dimensions[1].height = 22

        chart_dates = self.dates[-60:] if len(self.dates) > 60 else self.dates[:]
        offset = len(self.dates) - len(chart_dates)

        def _fill_block(start_row, title, metric, fmt):
            ws.cell(row=start_row, column=1, value=title)
            self._font(ws.cell(row=start_row, column=1), bold=True)
            ws.cell(row=start_row + 1, column=1, value="Date")
            for j, region in enumerate(REGION_ORDER):
                ws.cell(row=start_row + 1, column=2 + j, value=region)
                self._font(ws.cell(row=start_row + 1, column=2 + j), bold=True)
            for i, d in enumerate(chart_dates):
                ws.cell(row=start_row + 2 + i, column=1, value=d).number_format = "yyyy-mm-dd"
                for j, region in enumerate(REGION_ORDER):
                    agg = REGION_AGGREGATE[region]
                    v = self.value_at(agg, metric, offset + i)
                    if v is not None:
                        ws.cell(row=start_row + 2 + i, column=2 + j, value=v).number_format = fmt
            return start_row + 1, start_row + 1 + len(chart_dates)

        spr_h, spr_l = _fill_block(60, "Z-Spread by region — time series", METRICS["spread"], "0")
        nxt = spr_l + 5
        yld_h, yld_l = _fill_block(nxt, "Yield by region — time series", METRICS["yield"], "0.00")
        nxt2 = yld_l + 5
        # Rebased TR including the index
        ws.cell(row=nxt2, column=1, value="Cumulative Total Return rebased to 100")
        self._font(ws.cell(row=nxt2, column=1), bold=True)
        ws.cell(row=nxt2 + 1, column=1, value="Date")
        cols_to_plot = [INDEX_NAME] + [REGION_AGGREGATE[r] for r in REGION_ORDER]
        labels = ["EMBI Global"] + REGION_ORDER
        for j, lbl in enumerate(labels):
            ws.cell(row=nxt2 + 1, column=2 + j, value=lbl)
            self._font(ws.cell(row=nxt2 + 1, column=2 + j), bold=True)
        bases = {}
        for ent in cols_to_plot:
            for v in (self.series.get((ent, METRICS["tret"]), []) or [])[offset:]:
                if v is not None:
                    bases[ent] = v
                    break
            else:
                bases[ent] = None
        for i, d in enumerate(chart_dates):
            ws.cell(row=nxt2 + 2 + i, column=1, value=d).number_format = "yyyy-mm-dd"
            for j, ent in enumerate(cols_to_plot):
                v = self.value_at(ent, METRICS["tret"], offset + i)
                base = bases.get(ent)
                if v is not None and base:
                    ws.cell(row=nxt2 + 2 + i, column=2 + j, value=(v / base) * 100).number_format = "0.0"
        tr_h, tr_l = nxt2 + 1, nxt2 + 1 + len(chart_dates)

        def _make(title, ytitle, hdr, last, ncols):
            ch = LineChart()
            ch.title = title
            ch.style = 12
            ch.y_axis.title = ytitle
            ch.height = 10
            ch.width = 22
            data = Reference(ws, min_col=2, max_col=1 + ncols, min_row=hdr, max_row=last)
            cats = Reference(ws, min_col=1, min_row=hdr + 1, max_row=last)
            ch.add_data(data, titles_from_data=True)
            ch.set_categories(cats)
            return ch

        ws.add_chart(_make("Z-Spread by Region (bps)", "bps", spr_h, spr_l, len(REGION_ORDER)), "B3")
        ws.add_chart(_make("Yield to Maturity by Region (%)", "%", yld_h, yld_l, len(REGION_ORDER)), "B25")
        ws.add_chart(_make("Cumulative Total Return (rebased = 100)", "Index", tr_h, tr_l, len(cols_to_plot)), "B47")

    # --------- Data_Raw ----------
    def build_data_raw(self):
        ws = self.wb.create_sheet("Data_Raw")
        for i, h in enumerate(("Date", "Entity", "Metric", "Value"), start=1):
            self._hdr(ws.cell(row=1, column=i), h)
        ws.column_dimensions["A"].width = 12
        ws.column_dimensions["B"].width = 26
        ws.column_dimensions["C"].width = 18
        ws.column_dimensions["D"].width = 14

        keep = (
            [INDEX_NAME] + [REGION_AGGREGATE[r] for r in REGION_ORDER]
            + self.all_countries + self.ratings_present
        )
        row = 2
        for ent in keep:
            for metric in METRICS.values():
                vals = self.series.get((ent, metric))
                if not vals:
                    continue
                for d, v in zip(self.dates, vals):
                    if v is None:
                        continue
                    ws.cell(row=row, column=1, value=d).number_format = "yyyy-mm-dd"
                    ws.cell(row=row, column=2, value=ent)
                    ws.cell(row=row, column=3, value=metric)
                    ws.cell(row=row, column=4, value=v).number_format = "0.000000"
                    row += 1
        # Append weight history into Data_Raw too
        for c, vals in self.weight_series.items():
            for d, v in zip(self.weight_dates, vals):
                if v is None:
                    continue
                ws.cell(row=row, column=1, value=d).number_format = "yyyy-mm-dd"
                ws.cell(row=row, column=2, value=c)
                ws.cell(row=row, column=3, value="Index Weight (%)")
                ws.cell(row=row, column=4, value=v).number_format = "0.000000"
                row += 1

    # --------- orchestration ----------
    def build(self) -> Workbook:
        self.build_cover()
        self.build_instructions()
        self.build_forecast()
        self.build_weights()
        self._build_timeseries_sheet("Spreads", "spread", "bps", "0", "+0;-0;0")
        self._build_timeseries_sheet("Yields",  "yield",  "% YTM", "0.00", "+0.00;-0.00;0.00")
        self.build_tret_ytd()
        self.build_by_rating()
        self.build_snapshot()
        self.build_latam_focus()
        self.build_charts()
        self.build_weights_history()
        self.build_data_raw()
        return self.wb


# ============================================================================
# 6. CLI
# ============================================================================

def collect_inputs(args_inputs: List[str], recursive: bool) -> List[Path]:
    """Resolve user-provided paths to a list of CSV files."""
    out: List[Path] = []
    if not args_inputs:
        # Default: scan the cwd.
        base = Path.cwd()
        out.extend(sorted(base.rglob("*.csv") if recursive else base.glob("*.csv")))
        return out
    for raw in args_inputs:
        p = Path(raw).expanduser().resolve()
        if not p.exists():
            print(f"  WARN: not found, skipping: {p}", file=sys.stderr)
            continue
        if p.is_dir():
            out.extend(sorted(p.rglob("*.csv") if recursive else p.glob("*.csv")))
        else:
            out.append(p)
    return out


def main(argv: Optional[List[str]] = None) -> int:
    parser = argparse.ArgumentParser(description="Build EMBI Global dashboard from JPM CSVs.")
    parser.add_argument("inputs", nargs="*",
                        help="CSV files or directories (default: scan current directory).")
    parser.add_argument("--output", "-o", default="EMBI_Global_Dashboard.xlsx",
                        help="Output XLSX path (default: %(default)s).")
    parser.add_argument("--recursive", action="store_true",
                        help="Recursively scan directories for *.csv.")
    args = parser.parse_args(argv)

    paths = collect_inputs(args.inputs, args.recursive)
    if not paths:
        print("ERROR: no CSV files found", file=sys.stderr)
        return 1

    returns_results = []
    weights_results = []
    snapshot_results = []
    sources: List[Tuple[str, str]] = []
    skipped: List[str] = []

    for p in paths:
        kind = classify_csv(p)
        if kind == "returns":
            returns_results.append(load_returns(p))
            sources.append(("returns", p.name))
            print(f"  [returns]    {p.name}")
        elif kind == "weights_history":
            weights_results.append(load_weights_history(p))
            sources.append(("weights_history", p.name))
            print(f"  [weights]    {p.name}")
        elif kind == "snapshot":
            snapshot_results.append(load_snapshot(p))
            sources.append(("snapshot", p.name))
            print(f"  [snapshot]   {p.name}")
        else:
            skipped.append(p.name)

    if skipped:
        print(f"  Skipped {len(skipped)} non-JPM CSVs: {', '.join(skipped[:5])}{'...' if len(skipped) > 5 else ''}")

    if not returns_results:
        print("ERROR: at least one RETURNS file is required", file=sys.stderr)
        return 1

    dates, series = merge_returns(returns_results)
    weight_dates, weight_series = merge_weights_history(weights_results)
    snap_date, snap_data = merge_snapshots(snapshot_results)

    print(f"\n  Returns:  {len(dates)} dates × {len(series)} series")
    print(f"  Weights:  {len(weight_dates)} dates × {len(weight_series)} countries")
    print(f"  Snapshot: {snap_date.strftime('%Y-%m-%d') if snap_date else '(none)'} × {len(snap_data)} entities")

    builder = Builder(
        dates=dates, series=series,
        weight_dates=weight_dates, weight_series=weight_series,
        snap_date=snap_date, snap_data=snap_data,
        sources=sources,
    )
    wb = builder.build()
    out_path = Path(args.output).expanduser().resolve()
    wb.save(out_path)
    print(f"\n  Wrote {out_path}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
