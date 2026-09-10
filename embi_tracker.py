#!/usr/bin/env python3
"""
embi_tracker.py — EMBI Global Diversified: history, attribution and forecast.

WHAT IT DOES
------------
1. READS your historical workbook (EMBI_global_history.xlsx) and NEVER writes to
   it. That file is treated as read-only, always. Keep it in the folder.
2. READS every JPM sub-index download you drop in the folder ("JP latest.csv" or
   whatever it is called - the filename is irrelevant, the format is detected)
   and appends each one as a new dated row.
3. WRITES a growing copy, EMBI_history_extended.csv, in exactly the same shape
   as your workbook: Date plus "EM Debt Indices | <entity> | <metric>" columns.
   Your original stays untouched; this is the file that accumulates.
4. WRITES the dashboard, EMBI_Dashboard.xlsx, rebuilt from scratch every run.

Run it as often as you like. Re-running with no new files changes nothing.

    python embi_tracker.py                    # scan ., ingest, rebuild
    python embi_tracker.py --dir "C:/path"
    python embi_tracker.py --no-build         # ingest only
    python embi_tracker.py --check            # reconciliation report, no output file

WHY THE FILES ARE SPLIT THE WAY THEY ARE
----------------------------------------
Your history carries four metrics per entity: total-return index, yield, STW and
weight. The JPM download carries those plus durations and JPM's own return
attribution. Rather than bolt new columns onto your schema, the extras live in
EMBI_snapshot_extras.csv keyed by date. Your file shape stays exactly as JPM
ships it, and nothing downstream breaks when JPM adds a column.

A NOTE ON DURATION, WHICH MATTERS FOR EVERY ATTRIBUTION NUMBER HERE
-------------------------------------------------------------------
The history file has no duration column, and duration cannot be estimated
reliably from it: regressing daily returns on yield changes gives 3.3 for the
index against JPM's published 6.22, because the index YTM is contaminated by
distressed names whose quoted yields are arithmetic artefacts (Venezuela 42.8%,
Lebanon 80.3%, Ethiopia -1.0%). So attribution uses JPM's PUBLISHED spread
duration, taken from the nearest snapshot, and every such figure is labelled
"at published duration". Where no snapshot exists, DEFAULT_DURATION is used and
the workbook says so.
"""

from __future__ import annotations

import argparse
import csv
import math
import statistics
import sys
from collections import defaultdict
from datetime import datetime, timedelta
from pathlib import Path
from typing import Any, Dict, List, Optional, Sequence, Tuple

from openpyxl import Workbook, load_workbook
from openpyxl.chart import LineChart, Reference
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

# ===========================================================================
# CONFIGURATION
# ===========================================================================

HISTORY_XLSX_HINTS = ["embi_global_history", "embi global history", "history"]
EXTENDED_CSV = "EMBI_history_extended.csv"
EXTRAS_CSV = "EMBI_snapshot_extras.csv"
OUTPUT_XLSX = "EMBI_Dashboard.xlsx"
ARCHIVE_DIR = "snapshots_archive"

INDEX = "EMBIG Div"
M_TRI, M_YLD, M_STW, M_WGT = "Cum Tot Ret Idx", "Yld to Maturity", "STW(Trsy)", "Index Weight (%)"
METRICS = [M_TRI, M_YLD, M_STW, M_WGT]

REGIONS = ["LATIN Region", "EUROPE Region", "AFRICA Region", "MIDEAST Region", "ASIA Region"]
RATINGS = ["Credit AA only", "Credit A only", "Credit BBB only", "Credit BB only",
           "Credit B only", "Credit C only", "Credit IG only", "Credit Non-IG"]

# Used only when no JPM snapshot is available for a date. JPM's published
# EMBIGD spread duration on 08-Sep-2026.
DEFAULT_DURATION = 6.219
DEFAULT_IR_DURATION = 6.230

# Forecast betas: bp of spread per unit of driver. Practitioner priors, not
# regression estimates - documented on the Methodology tab with their known
# weakness (beta on rates is unconditional and cannot tell a growth-driven
# rate rise from an inflation shock).
BETAS = {"ust10y_bp": 0.30, "dxy_pct": 6.0, "vix_pts": 4.0, "rating_notches": -50.0}
CASH_RATE = 4.50           # the hurdle the call is judged against
# Scenarios are expressed as LEVELS of the US 10-year, not as changes: nobody
# forecasts "+25bp", they forecast "4.75%". Set the current level here; leave
# it None and the script falls back to the Treasury implied by the index
# itself (yield less spread), which sits at the index's ~11y point rather than
# the 10y and is therefore only a stand-in. Set it properly for published work.
UST10Y_NOW: Optional[float] = 4.80

# TERM PREMIUM
# ------------
# The index is NOT priced off the 10y. Its spread is struck at each bond's own
# maturity, and the cap-weighted average life is 10.2 years, so the Treasury it
# is actually exposed to sits further out: yield 6.65% less spread 173bp puts
# the anchor at 4.92% against a 10y of 4.80%. That 12bp is term premium for the
# extra maturity, and it is about what the cross-sectional curve implies
# (4.9bp per year of average life x ~1.2 years).
#
# The premium is computed each run, not hard-coded. What it buys is honesty on
# the scenario slide: a 10y at 5.00% means an index anchor at 5.12%, and the
# reader can see the two are different points on the curve.
#
# CURVE_BETA is how far the index's anchor moves for 1bp on the 10y. At 1.0 the
# curve shifts in parallel and the term premium cancels out of every CHANGE, so
# the returns are unaffected and the adjustment is presentational. Set it below
# 1.0 for a flattening (the long end moves less than the 10y) or above for a
# steepening. It cannot be estimated from these files - there is no 10y series
# in them - so it stays at 1.0 until you feed one in.
CURVE_BETA = 1.0
SENSITIVITY_UST = [-100, -75, -50, -25, 0, 25, 50, 75, 100]   # kept for grid B
SENSITIVITY_SPREADS = [125, 150, 175, 200, 225, 250, 275]
SCENARIO_SPREADS = [125, 150, 175, 200, 225, 250, 275, 300]   # rows of the scenario grid
# Bear / base / bull defaults for the Forecast tab. Every one of these is an input cell
# in the workbook; change them there, not here. "spread" is the level you actually want
# to argue (blank it in the workbook to let the drivers set it), "prob" the weight.
SCENARIOS = [
    ("Bear", dict(ust=50,  dxy=5.0,  vix=8.0,  rating=-0.25, spread=250, prob=25)),
    ("Base", dict(ust=-20, dxy=0.0,  vix=0.0,  rating=0.0,   spread=175, prob=50)),
    ("Bull", dict(ust=-50, dxy=-3.0, vix=-4.0, rating=0.25,  spread=150, prob=25)),
]
DISTRESSED_YIELD = 15.0     # above this a yield is a recovery bet, not a carry estimate
MAX_BREAK_DAYS = 24         # more break days than this and the name trades on price, not spread
SCN_COL = {"Bear": "I", "Base": "J", "Bull": "K"}      # fixed columns on the Forecast tab
SCN_ROW_UST, SCN_ROW_CHG = 6, 13                       # fixed rows: UST change, spread change
GRID_MARGIN_PP = 1.0        # within this of cash the cell is "marginal" (grey), not ok / not ok

# A duration approximation is a first-order expansion: it is accurate for small
# moves and fails badly for large ones, because duration itself changes as the
# credit moves. Beyond this many bp the attribution is flagged rather than
# quietly presented. Credit C moved -1059bp this year: -D x ds implies +52%
# against an actual +18%, so the legs there are arithmetic, not analysis.
LINEAR_LIMIT_BP = 200.0

# ---------------------------------------------------------------------------
# BASIS BREAKS
# ---------------------------------------------------------------------------
# JPM changes the index from time to time, and when they do the spread series
# steps on one day without the market having moved. The signature is a large
# jump in spread with no matching move in the total return index.
#
# 04-Sep-2026: JPM EXCLUDED DEFAULTED NAMES FROM THE SPREAD CALCULATION. They
# had an outsized influence on the aggregate - a handful of non-performing
# credits quoting four- and five-figure spreads were dragging the index number
# around. Confirmed by Alberto; the data agrees exactly. The index fell 63bp,
# which at duration 6.22 should have been worth +3.9%, while the index returned
# +0.1%. Only distressed paper moved (Lebanon x0.57, Venezuela x0.65,
# Credit C x0.32); performing credits were untouched (Mexico x0.99, Brazil
# x1.00, Egypt x1.00, IG x1.01).
#
# The important consequence: RETURNS are unaffected and remain comparable
# across the break - the bonds stay in the index for performance purposes, and
# the total return index is continuous. SPREADS before and after are not the
# same measure. So total return needs no adjustment and spread changes must be
# chain-linked, which is exactly what this script does.
#
# Every period change in this workbook is CHAIN-LINKED across such days: the
# jump itself is excluded, so what is reported is the market move. Both the
# chain-linked and the naive figure are shown wherever they differ, because the
# naive one is what anybody eyeballing two levels will compute.
BREAK_MIN_BP = 25.0        # candidate jump size
BREAK_MIN_GAP_PP = 1.0     # gap between actual and duration-implied return

# JPM snapshot -> history schema
SNAP_METRIC_MAP = {
    "Index Level": M_TRI,
    "Dur Wgt Yield Wrst": M_YLD, "Yield to Worst": M_YLD,
    "Dur Wgt STW (Trsy)": M_STW, "STW (Trsy)": M_STW,
    "Mkt Cap %": M_WGT,
}
SNAP_EXTRA_FIELDS = [
    "Spread Duration", "IR Duration to Worst", "Weighted Avg Avg Life Wrst",
    "IR Convexity to Worst", "Dur Wgt Z- Spread to Wrst",
    "YTD Change (%)", "Spread Return YTD Change (%)", "UST Return YTD Change (%)",
    "Coupon Return YTD Change (%)", "Price Return YTD Change (%)",
    "MTD Change (%)", "No. of Issues", "No. of Issuer",
    "Average S&P Rating", "Average Moody Rating", "Average Fitch Rating",
]
# Only DIVERSIFIED is accepted. Plain "EMBI Global" is a different index with
# different weights and a different level - Latin prints 939 on it against
# 1084 on Diversified for the same day - and letting one in produces a fake
# +15.7% "return". It is rejected at the gate, never mapped.
ACCEPTED_INDEX_LABELS = {"embi global diversified", "embig div"}
SNAP_ENTITY_MAP = {
    "embi global diversified": INDEX, "embig div": INDEX,
    "africa": "AFRICA Region", "asia": "ASIA Region", "europe": "EUROPE Region",
    "latin": "LATIN Region", "middle east": "MIDEAST Region",
    "investment grade": "Credit IG only", "non investment grade": "Credit Non-IG",
    "aa": "Credit AA only", "a": "Credit A only", "bbb": "Credit BBB only",
    "bb": "Credit BB only", "b": "Credit B only", "c": "Credit C only",
    "cote d'ivoire": "Cote d Ivoire", "trinidad and tobago": "Trinidad & Tobago",
}
SKIP_INSTRUMENTS = {"non latin", "sovereign", "quasi", "nr"}

# ===========================================================================
# STYLE
# ===========================================================================

FN = "Arial"
C_HDR, C_HDRFG = "002B5C", "FFFFFF"      # UBS-style deep blue header
C_BAND = "DCE3EC"                         # pale blue section band
C_GOOD, C_BAD = "EDF2F8", "FBE7E9"        # pale blue / pale red, no green
C_OK, C_MARG, C_NOTOK = "C9D9EE", "E3E3E3", "F2C4C9"   # scenario grid: ok / marginal / not ok
C_OKTXT, C_NOTOKTXT = "002B5C", "9C0A1E"
C_IN, C_BRD = "E8EEF5", "BFBFBF"          # inputs shaded blue, not yellow
C_HARD, C_TXT, C_NOTE = "0033A0", "000000", "6E6E6E"
BRD = Border(*[Side(style="thin", color=C_BRD)] * 4)


def F(c, size=10, bold=False, italic=False, color=C_TXT):
    c.font = Font(name=FN, size=size, bold=bold, italic=italic, color=color)
    return c


def H(c, txt):
    c.value = txt
    c.fill = PatternFill("solid", fgColor=C_HDR)
    c.font = Font(name=FN, size=10, bold=True, color=C_HDRFG)
    c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    c.border = BRD
    return c


def V(c, value, fmt="0.00", color=C_TXT, bold=False, fill=None):
    c.value = "" if value is None else value
    c.number_format = fmt
    c.border = BRD
    if fill:
        c.fill = PatternFill("solid", fgColor=fill)
    return F(c, bold=bold, color=color)


def L(c, txt, fill=None, bold=True, indent=0):
    c.value = txt
    c.border = BRD
    if fill:
        c.fill = PatternFill("solid", fgColor=fill)
    if indent:
        c.alignment = Alignment(indent=indent)
    return F(c, bold=bold)


def band(ws, row, ncols, txt):
    L(ws.cell(row=row, column=1), txt, fill=C_BAND)
    for c in range(2, ncols + 1):
        ws.cell(row=row, column=c).fill = PatternFill("solid", fgColor=C_BAND)
        ws.cell(row=row, column=c).border = BRD
    return row + 1


# ===========================================================================
# PARSING HELPERS
# ===========================================================================

def num(x) -> Optional[float]:
    if x is None:
        return None
    if isinstance(x, (int, float)):
        return None if (isinstance(x, float) and (math.isnan(x) or math.isinf(x))) else float(x)
    s = str(x).strip().replace(",", "").replace("%", "")
    if not s or s.upper() in {"NA", "N/A", "-", "--", "#N/A", "NULL", "NONE"}:
        return None
    if s.startswith("(") and s.endswith(")"):
        s = "-" + s[1:-1]
    try:
        v = float(s)
    except ValueError:
        return None
    return None if math.isnan(v) or math.isinf(v) else v


def to_date(x) -> Optional[str]:
    if isinstance(x, datetime):
        return x.strftime("%Y-%m-%d")
    s = str(x or "").strip()
    if not s:
        return None
    for f in ("%Y-%m-%d", "%d-%b-%Y", "%m/%d/%Y", "%d/%m/%Y", "%Y-%m-%d %H:%M:%S", "%d-%b-%y"):
        try:
            return datetime.strptime(s, f).strftime("%Y-%m-%d")
        except ValueError:
            continue
    return None


def dt(s: str) -> datetime:
    return datetime.strptime(s, "%Y-%m-%d")


def col_name(entity: str, metric: str) -> str:
    return f"EM Debt Indices | {entity} | {metric}"


def parse_col(h: str) -> Optional[Tuple[str, str]]:
    parts = [p.strip() for p in str(h).split("|")]
    return (parts[1], parts[2]) if len(parts) >= 3 else None


# ===========================================================================
# THE PANEL
# ===========================================================================

class Panel:
    """date -> {(entity, metric): value}. Wide on disk, dict in memory."""

    def __init__(self) -> None:
        self.data: Dict[str, Dict[Tuple[str, str], float]] = {}
        self.extras: Dict[str, Dict[Tuple[str, str], Any]] = {}
        self.sources: Dict[str, str] = {}

    # ---- load ----
    def load_history_xlsx(self, path: Path) -> int:
        """READ ONLY. This file is never written to."""
        wb = load_workbook(path, read_only=True, data_only=True)
        ws = wb[wb.sheetnames[0]]
        it = ws.iter_rows(values_only=True)
        hdr = next(it)
        cols = {j: parse_col(h) for j, h in enumerate(hdr) if j and h}
        n = 0
        for row in it:
            d = to_date(row[0])
            if not d:
                continue
            bucket = self.data.setdefault(d, {})
            for j, key in cols.items():
                if key is None or j >= len(row):
                    continue
                v = num(row[j])
                if v is not None:
                    bucket[key] = v
            self.sources.setdefault(d, "history workbook")
            n += 1
        wb.close()
        return n

    def load_extended_csv(self, path: Path) -> int:
        if not path.exists():
            return 0
        n = 0
        with path.open("r", newline="", encoding="utf-8-sig") as f:
            rd = csv.reader(f)
            hdr = next(rd, None)
            if not hdr:
                return 0
            cols = {j: parse_col(h) for j, h in enumerate(hdr) if j and h}
            for row in rd:
                d = to_date(row[0]) if row else None
                if not d:
                    continue
                bucket = self.data.setdefault(d, {})
                for j, key in cols.items():
                    if key is None or j >= len(row):
                        continue
                    v = num(row[j])
                    if v is not None:
                        bucket.setdefault(key, v)
                self.sources.setdefault(d, "extended csv")
                n += 1
        return n

    def load_extras_csv(self, path: Path) -> None:
        if not path.exists():
            return
        with path.open("r", newline="", encoding="utf-8-sig") as f:
            for r in csv.DictReader(f):
                d = to_date(r.get("date"))
                if not d:
                    continue
                self.extras.setdefault(d, {})[(r.get("entity", ""), r.get("field", ""))] = r.get("value")

    # ---- ingest a JPM download ----
    @staticmethod
    def is_snapshot(path: Path) -> bool:
        try:
            with path.open("r", newline="", encoding="utf-8-sig", errors="replace") as f:
                for row in csv.reader(f):
                    if any((c or "").strip() for c in row):
                        return (row[0] or "").strip().strip('"').lower() in {"bam id", "bamid", "bam_id"}
        except OSError:
            return False
        return False

    def ingest_snapshot(self, path: Path, known_entities: Sequence[str]) -> Optional[Tuple[str, int, int]]:
        with path.open("r", newline="", encoding="utf-8-sig", errors="replace") as f:
            rows = [{(k or "").strip(): v for k, v in r.items()} for r in csv.DictReader(f)]
        lookup = {e.lower(): e for e in known_entities}
        date = None
        vals: Dict[Tuple[str, str], float] = {}
        ext: Dict[Tuple[str, str], Any] = {}
        first = True
        for r in rows:
            bam = (r.get("Bam Id") or "").strip()
            inst = (r.get("Instrument") or "").strip()
            if bam.lower().startswith(("disclaimer", "http")):
                break
            if inst.startswith("By ") or not inst:
                continue
            if date is None:
                date = to_date(r.get("Date"))
            key = inst.lower()
            if first:
                first = False
                # ---- THE GATE ----
                if key not in ACCEPTED_INDEX_LABELS:
                    print(f"REJECTED : {path.name} is '{inst}', not EMBI Global Diversified. "
                          f"Not ingested. Remove it from the folder.", file=sys.stderr)
                    return None
                # Same date already in the history? The index level must agree.
                # If it does not, this file is a different index wearing the
                # right label, and it must not overwrite anything.
                lvl = num(r.get("Index Level"))
                have = self.data.get(date or "", {}).get((INDEX, M_TRI))
                if lvl is not None and have is not None and abs(lvl / have - 1.0) > 0.005:
                    print(f"REJECTED : {path.name} index level {lvl:.2f} disagrees with the "
                          f"history's {have:.2f} on {date} by {abs(lvl / have - 1) * 100:.1f}%. "
                          f"Different index. Not ingested.", file=sys.stderr)
                    return None
                ent = INDEX
            elif key in SNAP_ENTITY_MAP:
                ent = SNAP_ENTITY_MAP[key]
            elif key in SKIP_INSTRUMENTS:
                continue
            elif key in lookup:
                ent = lookup[key]
            else:
                ent = inst           # a new country JPM has added
            for fld, met in SNAP_METRIC_MAP.items():
                v = num(r.get(fld))
                if v is not None:
                    vals[(ent, met)] = v
            for fld in SNAP_EXTRA_FIELDS:
                raw = r.get(fld)
                if raw not in (None, ""):
                    ext[(ent, fld)] = str(raw).strip()
        if not date or not vals:
            return None
        before = len(self.data.get(date, {}))
        self.data.setdefault(date, {}).update(vals)
        self.extras.setdefault(date, {}).update(ext)
        self.sources[date] = f"JPM snapshot ({path.name})"
        return date, len(vals), len(self.data[date]) - before

    # ---- write ----
    def entities(self) -> List[str]:
        seen: Dict[str, None] = {}
        for b in self.data.values():
            for (e, _m) in b:
                seen.setdefault(e, None)
        out = [INDEX]
        out += [e for e in REGIONS if e in seen]
        out += [e for e in RATINGS if e in seen]
        out += sorted(e for e in seen if e not in out)
        return out

    def dates(self) -> List[str]:
        return sorted(self.data)

    def save_extended(self, path: Path) -> None:
        ents = self.entities()
        cols = [(e, m) for e in ents for m in METRICS]
        with path.open("w", newline="", encoding="utf-8") as f:
            w = csv.writer(f)
            w.writerow(["Date"] + [col_name(e, m) for e, m in cols])
            for d in self.dates():
                b = self.data[d]
                w.writerow([d] + ["" if b.get(k) is None else b[k] for k in cols])

    def save_extras(self, path: Path) -> None:
        with path.open("w", newline="", encoding="utf-8") as f:
            w = csv.writer(f)
            w.writerow(["date", "entity", "field", "value"])
            for d in sorted(self.extras):
                for (e, fld), v in sorted(self.extras[d].items()):
                    w.writerow([d, e, fld, v])

    # ---- read ----
    def get(self, d: str, e: str, m: str) -> Optional[float]:
        return self.data.get(d, {}).get((e, m))

    def extra(self, d: str, e: str, fld: str) -> Optional[str]:
        return self.extras.get(d, {}).get((e, fld))

    def series(self, e: str, m: str) -> List[Tuple[str, float]]:
        out = []
        for d in self.dates():
            v = self.get(d, e, m)
            if v is not None:
                out.append((d, v))
        return out

    def nearest_extra(self, d: str, e: str, fld: str) -> Optional[float]:
        """Published value from the closest snapshot date."""
        cands = [(abs((dt(k) - dt(d)).days), k) for k in self.extras
                 if (e, fld) in self.extras[k]]
        if not cands:
            return None
        cands.sort()
        return num(self.extras[cands[0][1]][(e, fld)])

    def duration(self, d: str, e: str) -> Tuple[float, str]:
        v = self.nearest_extra(d, e, "Spread Duration")
        if v:
            return v, "published"
        return DEFAULT_DURATION, "default"

    def ir_duration(self, d: str, e: str) -> float:
        return self.nearest_extra(d, e, "IR Duration to Worst") or DEFAULT_IR_DURATION

    def spread_meaningful(self, e: str) -> bool:
        """False for names whose spread series steps constantly — defaulted or deeply
        distressed paper that trades on price. Their spread changes are not reported."""
        return len(self.breaks_for(e)) <= MAX_BREAK_DAYS

    def breaks_for(self, e: str) -> Dict[str, Dict[str, Any]]:
        """Basis-break days for this sub-index, detected on its own series and cached."""
        cache = self.__dict__.setdefault("_breaks_by_entity", {})
        if e not in cache:
            cache[e] = detect_breaks(self, e)
        return cache[e]


# ===========================================================================
# ANALYTICS
# ===========================================================================

def detect_breaks(panel: "Panel", e: str = INDEX) -> Dict[str, Dict[str, Any]]:
    """Days where a sub-index's spread (and yield) series steps but its return series
    does not — a composition or methodology change, not the market. Run per entity:
    the Middle East steps at month-ends as Lebanon's weight is rebalanced, and the
    index-level detector never sees it.

    A day is a break when the return the yield move implies (minus duration times
    the move) is more than BREAK_MIN_GAP_PP away from the actual return AND the gap
    is most of the implied move, so a distressed name's convexity noise on a real
    move is not mistaken for a break."""
    out: Dict[str, Dict[str, Any]] = {}
    ds = panel.dates()
    dur, _src = panel.duration(ds[-1], e) if ds else (DEFAULT_DURATION, "default")
    for i in range(1, len(ds)):
        d0, d1 = ds[i - 1], ds[i]
        s0, s1 = panel.get(d0, e, M_STW), panel.get(d1, e, M_STW)
        t0, t1 = panel.get(d0, e, M_TRI), panel.get(d1, e, M_TRI)
        if None in (s0, s1, t0, t1) or not t0:
            continue
        jump = s1 - s0
        if abs(jump) < BREAK_MIN_BP:
            continue
        actual = (t1 / t0 - 1.0) * 100.0
        # Test the return against the YIELD move, not the spread move: on a day when
        # spreads widen 29bp while Treasuries rally 22bp (28-Sep-2022) the return is
        # flat for a market reason, and calling that a basis break would drop a real
        # move from every window that crosses it.
        y0, y1 = panel.get(d0, e, M_YLD), panel.get(d1, e, M_YLD)
        move = (y1 - y0) * 100.0 if None not in (y0, y1) else jump
        implied = -dur * move / 100.0
        gap = abs(actual - implied)
        if gap > BREAK_MIN_GAP_PP and gap > 0.6 * abs(implied):
            out[d1] = {"jump_bp": jump, "actual_pct": actual, "implied_pct": implied}
    return out


def chain_change(panel: "Panel", e: str, metric: str, d0: Optional[str], d1: str,
                 breaks: Dict[str, Any]) -> Tuple[Optional[float], Optional[float], bool]:
    """(chain-linked change, naive change, crossed a break?).

    Chain-linked sums the day-to-day moves and skips break days, so it measures
    the market. Naive is simply end minus start."""
    if not d0:
        return None, None, False
    v0, v1 = panel.get(d0, e, metric), panel.get(d1, e, metric)
    scale = 100.0 if metric == M_YLD else 1.0
    naive = None
    if v0 is not None and v1 is not None:
        naive = (v1 / v0 - 1.0) * 100.0 if metric == M_TRI else (v1 - v0) * scale
    breaks = panel.breaks_for(e)          # the entity's own break days, not the index's
    window = [d for d in panel.dates() if d0 <= d <= d1]
    crossed = any(d in breaks for d in window[1:])
    if metric == M_TRI or not crossed:
        return naive, naive, crossed
    # Carry the last good observation forward rather than skipping a delta:
    # dropping both sides of a gap loses whatever moved across it. The history
    # has holes (2026-04-03 for one) and skipping cost 5bp on the Middle East.
    tot = 0.0
    prev = None
    for d in window:
        cur = panel.get(d, e, metric)
        if cur is None:
            continue
        if prev is not None and d not in breaks:
            tot += (cur - prev) * scale
        prev = cur
    return tot, naive, True


def year_start(panel: Panel, year: int) -> Optional[str]:
    """Last observation of the prior year - the correct YTD denominator."""
    prior = [d for d in panel.dates() if dt(d).year < year]
    return prior[-1] if prior else None


def implied_ust(panel: Panel, d: str, e: str) -> Optional[float]:
    """Yield less spread = the Treasury the entity is actually struck against."""
    y = panel.get(d, e, M_YLD)
    s = panel.get(d, e, M_STW)
    return None if y is None or s is None else y - s / 100.0


def attribute(panel: Panel, e: str, d0: str, d1: str) -> Optional[Dict[str, Any]]:
    """Total return decomposed into a spread leg, a rate leg and the remainder.

    The identity holds by construction:
        total = -D x d(spread) - D x d(UST) + (carry & residual)
    The remainder is not an error term to be minimised - it is realised coupon,
    roll, rebalancing and defaults, and for this index it runs far below the
    quoted YTM because distressed names carry yields they never pay.
    """
    t0, t1 = panel.get(d0, e, M_TRI), panel.get(d1, e, M_TRI)
    s0, s1 = panel.get(d0, e, M_STW), panel.get(d1, e, M_STW)
    if None in (t0, t1) or not t0:
        return None
    total = (t1 / t0 - 1.0) * 100.0
    out: Dict[str, Any] = {"total": total, "spread_bp": None, "ust_bp": None,
                           "spread_leg": None, "rate_leg": None, "residual": None}
    if s0 is None or s1 is None:
        return out
    u0, u1 = implied_ust(panel, d0, e), implied_ust(panel, d1, e)
    D, src = panel.duration(d1, e)
    breaks = getattr(panel, "_breaks", {}) or {}
    chained, naive, crossed = chain_change(panel, e, M_STW, d0, d1, breaks)
    move = chained if chained is not None else (s1 - s0)
    if not panel.spread_meaningful(e):
        out["spread_bp"] = None
        out["spread_bp_naive"] = naive
        out["crossed_break"] = crossed
        out["dur"], out["dur_src"] = D, src
        out["flag"] = "spread not meaningful — trades on price (defaulted / distressed)"
        return out
    out["spread_bp"] = move
    out["spread_bp_naive"] = naive
    out["crossed_break"] = crossed
    out["dur"], out["dur_src"] = D, src
    out["spread_leg"] = -D * move / 100.0
    if u0 is not None and u1 is not None:
        # The 04-Sep-2026 break moved YIELD and SPREAD together (YTM 7.23 -> 6.63,
        # STW 236 -> 173), so their difference - the implied Treasury - barely
        # moved: 4.869% -> 4.902%. It is chain-linked anyway, using the same
        # break dates, rather than corrected by an assumption about which leg
        # absorbed the change.
        du = 0.0
        window = [d for d in panel.dates() if d0 <= d <= d1]
        prev_ = None
        for d in window:
            cur_ = implied_ust(panel, d, e)
            if cur_ is None:
                continue
            if prev_ is not None and d not in breaks:
                du += (cur_ - prev_) * 100.0
            prev_ = cur_
        if len(window) < 2:
            du = (u1 - u0) * 100.0
        out["ust_bp"] = du
        out["rate_leg"] = -D * du / 100.0
        out["residual"] = total - out["spread_leg"] - out["rate_leg"]
    # Is the decomposition worth reading? Two ways it stops being so: the move
    # is too large for a first-order expansion, or the residual is doing more
    # work than the total, which means the legs are cancelling rather than
    # explaining.
    # NOTE: do not flag on "residual is large relative to the total". The
    # residual is mostly carry, so for any entity whose year netted to roughly
    # zero it is legitimately several percent - that rule fires on the
    # healthiest rows in the book. What is actually suspicious is a residual
    # that is materially NEGATIVE, since coupon cannot be negative: that means
    # defaults, restructurings or index composition, not carry.
    reasons = []
    if crossed:
        reasons.append("chain-linked across a basis break")
    if abs(out["spread_bp"] or 0) > LINEAR_LIMIT_BP:
        reasons.append(f"spread moved {abs(out['spread_bp']):.0f}bp — too large for a duration split")
    if out["residual"] is not None and out["residual"] < -1.0:
        reasons.append("negative carry — defaults or composition, not coupon")
    # Carry cannot plausibly exceed the yield earned over the window. When the
    # plug does, the duration approximation is absorbing something it should
    # not and the legs are indicative at best.
    yld = panel.get(d1, e, M_YLD)
    if out["residual"] is not None and yld:
        yrs = max((dt(d1) - dt(d0)).days / 365.0, 1e-6)
        if out["residual"] > 1.25 * yld * yrs + 0.5:
            reasons.append("carry plug exceeds the yield — legs are indicative only")
    w1 = panel.get(d1, e, M_WGT)
    w0 = panel.get(d0, e, M_WGT)
    if w0 and w1 and abs(w1 - w0) > max(1.0, 0.5 * w0):
        reasons.append("index weight changed materially")
    out["flag"] = "; ".join(reasons)
    return out


def sigma_12m(panel: Panel, e: str = INDEX) -> Tuple[Optional[float], Optional[float], int]:
    """Two readings of 12m spread vol: daily scaled by root-252, and the actual
    distribution of overlapping 252-day changes. They disagree when spreads
    mean-revert, and the second is the honest one for a 12m horizon."""
    ser = panel.series(e, M_STW)
    if len(ser) < 30:
        return None, None, 0
    diffs = [ser[i][1] - ser[i - 1][1] for i in range(1, len(ser))]
    daily = statistics.pstdev(diffs) * math.sqrt(252)
    obs = [ser[i][1] - ser[i - 252][1] for i in range(252, len(ser))]
    direct = statistics.pstdev(obs) if len(obs) >= 30 else None
    return daily, direct, len(obs)


def ust_levels(now: float, span: float = 1.0, step: float = 0.25) -> List[float]:
    """Quarter-point levels either side of where the 10y is today."""
    lo = math.floor((now - span) / step) * step
    hi = math.ceil((now + span) / step) * step
    out, x = [], lo
    while x <= hi + 1e-9:
        out.append(round(x, 4))
        x += step
    return out


def chained_series(panel: Panel, e: str, metric: str, breaks: Dict[str, Any]) -> Dict[str, float]:
    """The spread history restated on today's basis: start from today's published level
    and walk back with the chain-linked daily moves, so every basis-break step is removed.
    The steps are removed proportionally (each day's ratio, not its difference): the
    defaulted names that JPM took out on 04-Sep contributed far more basis points when
    spreads were wide in 2022 than they do today, and a proportional restatement respects
    that; a fixed-bp restatement drives some regional histories negative. Still an
    approximation — see break_share() for how much of a segment's level the steps were."""
    ser = panel.series(e, metric)
    breaks = panel.breaks_for(e)
    if not ser:
        return {}
    out = {ser[-1][0]: ser[-1][1]}
    lvl = ser[-1][1]
    for i in range(len(ser) - 1, 0, -1):
        d, v = ser[i]
        dp, vp = ser[i - 1]
        if d not in breaks and v and vp and v > 0 and vp > 0:
            lvl *= vp / v
        out[dp] = lvl
    return out


def break_share(panel: Panel, e: str, breaks: Dict[str, Any] = None) -> float:
    """How much of the segment's spread history has been restated: the sum over its
    break days of |step| / level before the step. Above ~0.5 the composition changed
    too much for the restated history to be more than indicative."""
    ser = panel.series(e, M_STW)
    brk = panel.breaks_for(e)
    tot = 0.0
    for i in range(1, len(ser)):
        if ser[i][0] in brk and ser[i - 1][1]:
            tot += abs(ser[i][1] - ser[i - 1][1]) / abs(ser[i - 1][1])
    return tot


def pctl(vals: Sequence[float], q: float) -> Optional[float]:
    """Linear-interpolated percentile, q in [0, 1]."""
    xs = sorted(v for v in vals if v is not None)
    if not xs:
        return None
    if len(xs) == 1:
        return xs[0]
    pos = q * (len(xs) - 1)
    lo, hi = int(math.floor(pos)), int(math.ceil(pos))
    return xs[lo] + (xs[hi] - xs[lo]) * (pos - lo)


def pct_rank(vals: Sequence[float], x: float) -> Optional[float]:
    xs = [v for v in vals if v is not None]
    if not xs:
        return None
    return 100.0 * sum(1 for v in xs if v <= x) / len(xs)


def rolling_12m_returns(panel: Panel, e: str = INDEX, obs: int = 252) -> List[Tuple[str, float]]:
    """(start date, total return % over the following ~12 months) for every start."""
    ser = panel.series(e, M_TRI)
    return [(ser[i - obs][0], (ser[i][1] / ser[i - obs][1] - 1.0) * 100.0)
            for i in range(obs, len(ser)) if ser[i - obs][1]]


def worst_drawdowns(panel: Panel, e: str = INDEX, n: int = 3) -> List[Dict[str, Any]]:
    """The n deepest peak-to-trough falls in the total return index, with recovery dates."""
    ser = panel.series(e, M_TRI)
    if len(ser) < 2:
        return []
    episodes: List[Dict[str, Any]] = []
    peak_d, peak_v = ser[0]
    trough_d, trough_v = ser[0]
    in_dd = False
    for d, v in ser[1:]:
        if v >= peak_v:
            if in_dd:
                episodes.append(dict(peak=peak_d, trough=trough_d, recovered=d,
                                     depth=(trough_v / peak_v - 1.0) * 100.0,
                                     days_down=(dt(trough_d) - dt(peak_d)).days,
                                     days_back=(dt(d) - dt(trough_d)).days))
                in_dd = False
            peak_d, peak_v = d, v
            trough_d, trough_v = d, v
        else:
            in_dd = True
            if v < trough_v:
                trough_d, trough_v = d, v
    if in_dd:
        episodes.append(dict(peak=peak_d, trough=trough_d, recovered=None,
                             depth=(trough_v / peak_v - 1.0) * 100.0,
                             days_down=(dt(trough_d) - dt(peak_d)).days, days_back=None))
    episodes.sort(key=lambda x: x["depth"])
    return episodes[:n]


def total_return(yld: float, ir_dur: float, spr_dur: float,
                 d_ust_bp: float, spread_now: float, spread_fcst: float) -> float:
    return yld - ir_dur * (d_ust_bp / 100.0) - spr_dur * ((spread_fcst - spread_now) / 100.0)


def scenario_grid(dash: "Dashboard"):
    """12m total return at every (spread level, US 10y level) pair. Spread and rates are set
    independently here — no beta — so the reader sees the raw arithmetic. Returns
    (levels, rows) where rows is a list of (label, spread_bp, [tr per level]). The first row
    is the spot spread so the reader can find today on the grid."""
    levels = list(dash.ust_grid)
    if all(abs(l - dash.ust10y) > 0.005 for l in levels):
        levels = sorted(levels + [dash.ust10y])       # the spot 10y gets its own column
    spreads = [("Spot %.0fbp" % dash.spread, dash.spread)] + \
              [("%dbp" % s, float(s)) for s in SCENARIO_SPREADS]
    rows = []
    for lab, s_ in spreads:
        vals = []
        for lvl in levels:
            u = (lvl - dash.ust10y) * 100.0
            anchor = lvl + dash.term_premium + (CURVE_BETA - 1.0) * u / 100.0
            du = (anchor - dash.ust_now) * 100.0
            vals.append(total_return(dash.yld, dash.IRD, dash.D, du, dash.spread, s_))
        rows.append((lab, s_, vals))
    return levels, rows


def grid_verdict(tr: float) -> str:
    """'ok' clears cash by GRID_MARGIN_PP or more, 'marginal' is within the margin either
    side of cash, 'bad' is below cash by more than the margin."""
    if tr >= CASH_RATE + GRID_MARGIN_PP:
        return "ok"
    if tr >= CASH_RATE - GRID_MARGIN_PP:
        return "marginal"
    return "bad"


# ===========================================================================
# WORKBOOK
# ===========================================================================

class Dashboard:
    def __init__(self, panel: Panel) -> None:
        self.p = panel
        self.dates = panel.dates()
        self.d1 = self.dates[-1]
        self.year = dt(self.d1).year
        self.d0 = year_start(panel, self.year) or self.dates[0]
        self.breaks = getattr(panel, "_breaks", {}) or {}
        self.wb = Workbook()
        self.wb.remove(self.wb.active)
        self.D, self.dsrc = panel.duration(self.d1, INDEX)
        self.IRD = panel.ir_duration(self.d1, INDEX)
        self.yld = panel.get(self.d1, INDEX, M_YLD) or 0.0
        self.spread = panel.get(self.d1, INDEX, M_STW) or 0.0
        self.countries = [e for e in panel.entities()
                          if e != INDEX and e not in REGIONS and e not in RATINGS]
        iu = implied_ust(panel, self.d1, INDEX)
        self.ust_now = iu if iu is not None else 4.50      # the index's own anchor
        self.ust_src = ("index anchor = yield less spread; the 10y is set in the script"
                        if UST10Y_NOW is not None else
                        "no 10y set — the index anchor is standing in for it")
        # self.ust_now is the INDEX's own Treasury anchor. The 10y is separate.
        self.ust10y = float(UST10Y_NOW) if UST10Y_NOW is not None else self.ust_now
        self.term_premium = self.ust_now - self.ust10y
        self.ust_grid = ust_levels(self.ust10y)

    # ---------- cover ----------
    def cover(self):
        ws = self.wb.create_sheet("Cover")
        ws.sheet_view.showGridLines = False
        ws.column_dimensions["A"].width = 40
        ws.column_dimensions["B"].width = 66
        ws["A1"] = "EMBI Global Diversified — history, attribution and forecast"
        F(ws["A1"], size=16, bold=True, color=C_HDR)
        a = attribute(self.p, INDEX, self.d0, self.d1) or {}
        rows = [
            ("Latest observation", self.d1),
            ("YTD measured from", f"{self.d0} (last print of {self.year - 1})"),
            ("Observations in history", f"{len(self.dates):,} from {self.dates[0]}"),
            ("Index spread (STW Trsy)", f"{self.spread:.0f}bp"),
            ("Index yield", f"{self.yld:.2f}%"),
            ("Spread duration", f"{self.D:.2f}  ({self.dsrc})"),
            ("", ""),
            ("YTD total return", f"{a.get('total', 0):+.2f}%"),
            ("YTD spread move", f"{a.get('spread_bp', 0):+.0f}bp"),
            ("YTD move in the underlying Treasury", f"{a.get('ust_bp', 0):+.0f}bp"),
            ("", ""),
            ("Countries tracked", len(self.countries)),
            ("Generated", datetime.now().strftime("%Y-%m-%d %H:%M")),
        ]
        r = 3
        for k, v in rows:
            if k:
                L(ws.cell(row=r, column=1), k, fill=C_BAND)
                V(ws.cell(row=r, column=2), v, fmt="General", color=C_HARD, bold=True)
            r += 1
        r += 1
        for line in [
            "Your history workbook is read-only and is never modified. Each run re-reads it,",
            "adds any JPM downloads found in the folder, and writes the growing copy to",
            f"{EXTENDED_CSV}. This dashboard is rebuilt from scratch every time - delete it",
            "and rerun and you lose nothing.",
            "",
            "Spread attribution uses JPM's published spread duration, not one estimated from",
            "the history: the index YTM is distorted by distressed names whose quoted yields",
            "are arithmetic artefacts, so a regression-based duration comes out near 3.3",
            "against a published 6.22. See the Methodology tab.",
        ]:
            ws.cell(row=r, column=1, value=line)
            F(ws.cell(row=r, column=1), italic=True, color=C_NOTE)
            r += 1

    # ---------- YTD dashboard ----------
    def ytd(self):
        ws = self.wb.create_sheet("YTD_Summary")
        ws.sheet_view.showGridLines = False
        ws["A1"] = f"Year to date {self.year} — {self.d0} to {self.d1}"
        F(ws["A1"], size=14, bold=True, color=C_HDR)
        ws["A2"] = ("Total return is (end / start - 1) on the index's own total-return series. "
                    "Spread and Treasury moves are shown in basis points. For a return "
                    "attribution use JPM's published legs below — ours has been removed because a "
                    "duration split misbehaves badly on this index.")
        F(ws["A2"], italic=True, color=C_NOTE)
        ws.column_dimensions["A"].width = 34
        for c in "BCDEFGH":
            ws.column_dimensions[c].width = 15

        a = attribute(self.p, INDEX, self.d0, self.d1) or {}
        r = 4
        r = band(ws, r, 3, "EMBI Global Diversified")
        for lbl, val, fmt, note in [
            ("Total return", a.get("total"), "+0.00;-0.00", ""),
            ("Spread move (bp)", a.get("spread_bp"), "+0;-0",
             f"{self.p.get(self.d0, INDEX, M_STW):.0f} -> {self.spread:.0f}"),
            ("Underlying Treasury move (bp)", a.get("ust_bp"), "+0;-0", "yield less spread"),
        ]:
            if not lbl:
                r += 1
                continue
            L(ws.cell(row=r, column=1), lbl, bold=not lbl.startswith("  "))
            V(ws.cell(row=r, column=2), val, fmt=fmt, color=C_HARD, bold=True)
            ws.cell(row=r, column=3, value=note)
            F(ws.cell(row=r, column=3), italic=True, color=C_NOTE)
            r += 1

        # JPM's own attribution, when a snapshot covers this date
        jt = self.p.nearest_extra(self.d1, INDEX, "YTD Change (%)")
        js = self.p.nearest_extra(self.d1, INDEX, "Spread Return YTD Change (%)")
        ju = self.p.nearest_extra(self.d1, INDEX, "UST Return YTD Change (%)")
        if jt is not None:
            r += 1
            r = band(ws, r, 3, "JPM's own published attribution — cite this one externally")
            for lbl, v in [("Total return", jt), ("Spread return", js), ("UST return", ju)]:
                L(ws.cell(row=r, column=1), lbl, bold=False)
                V(ws.cell(row=r, column=2), v, fmt="+0.00;-0.00", color=C_HARD)
                r += 1
            L(ws.cell(row=r, column=1), "  our total vs JPM's", bold=False)
            diff = (a.get("total") or 0) - jt
            V(ws.cell(row=r, column=2), diff, fmt="+0.000;-0.000",
              fill=C_GOOD if abs(diff) < 0.02 else C_BAD)
            ws.cell(row=r, column=3, value="reconciliation check on the total-return index")
            F(ws.cell(row=r, column=3), italic=True, color=C_NOTE)
            r += 2
            ws.cell(row=r, column=1, value=("JPM's legs and ours differ: theirs are computed on the "
                                            "actual Treasury hedge and compound through the year, "
                                            "ours are a duration approximation at a single duration. "
                                            "Use JPM's in client work; ours travels across every "
                                            "date in the history, theirs does not."))
            F(ws.cell(row=r, column=1), italic=True, color=C_NOTE)
            r += 2

        # movers
        r += 1
        rows = []
        for e in self.countries:
            at = attribute(self.p, e, self.d0, self.d1)
            w = self.p.get(self.d1, e, M_WGT)
            if at and at["total"] is not None and w:
                rows.append((at["total"], e, w, at.get("spread_bp")))
        rows.sort(reverse=True)
        for title, sel in [("Best YTD total return", rows[:8]),
                           ("Worst YTD total return", rows[-8:][::-1])]:
            r = band(ws, r, 4, title)
            H(ws.cell(row=r, column=2), "YTD return %")
            H(ws.cell(row=r, column=3), "weight %")
            H(ws.cell(row=r, column=4), "spread move bp")
            r += 1
            for tot, e, w, sb in sel:
                L(ws.cell(row=r, column=1), e, bold=False)
                V(ws.cell(row=r, column=2), tot, fmt="+0.00;-0.00", color=C_HARD)
                V(ws.cell(row=r, column=3), w, fmt="0.00")
                V(ws.cell(row=r, column=4), sb, fmt="+0;-0")
                r += 1
            r += 1

    # ---------- breakdown tabs ----------
    def _breakdown(self, name: str, title: str, entities: List[str], sort_by_weight: bool):
        ws = self.wb.create_sheet(name)
        ws.sheet_view.showGridLines = False
        ws["A1"] = f"{title} — YTD {self.year} ({self.d0} to {self.d1})"
        F(ws["A1"], size=14, bold=True, color=C_HDR)
        ws["A2"] = ("Every return here is (end / start - 1) on that sub-index's own total-return "
                    "index. Nothing is derived from duration. Spread and Treasury moves are shown "
                    "in basis points beside them; the duration decomposition has been removed "
                    "because it produced figures like a +52% spread contribution for Credit C "
                    "against an actual +17.9%.")
        F(ws["A2"], italic=True, color=C_NOTE)
        d_1m, d_3m, d_12m = self._back(1), self._back(3), self._back(12)
        cols = [("", 30), ("Weight %", 10), ("Spread now", 11), ("Yield %", 9), ("Duration", 10),
                (f"Return 1m %\n{d_1m or '—'}", 13),
                (f"Return 3m %\n{d_3m or '—'}", 13),
                (f"Return YTD %\n{self.d0}", 13),
                (f"Return 12m %\n{d_12m or '—'}", 13),
                ("Spread 1m bp", 12), ("Spread YTD bp", 13), ("Treasury YTD bp", 14),
                ("read with care", 34)]
        hr = 4
        for i, (h, w) in enumerate(cols, start=1):
            H(ws.cell(row=hr, column=i), h)
            ws.column_dimensions[get_column_letter(i)].width = w
        rows = []
        for e in entities:
            at = attribute(self.p, e, self.d0, self.d1)
            if not at:
                continue
            w = self.p.get(self.d1, e, M_WGT) or 0.0
            if w <= 0:
                continue          # not in the index at the latest date
            rows.append((w if sort_by_weight else 0, e, at, w))
        rows.sort(key=lambda t: -t[0])
        ws.row_dimensions[hr].height = 30
        r = hr + 1
        for _k, e, at, w in rows:
            L(ws.cell(row=r, column=1), e, bold=False)
            V(ws.cell(row=r, column=2), w, fmt="0.00", color=C_HARD)
            V(ws.cell(row=r, column=3), self.p.get(self.d1, e, M_STW), fmt="0", color=C_HARD)
            V(ws.cell(row=r, column=4), self.p.get(self.d1, e, M_YLD), fmt="0.00", color=C_HARD)
            V(ws.cell(row=r, column=5), at.get("dur"), fmt="0.00")
            # total return = (end / start - 1) on this sub-index's own return series
            for k, dd in enumerate([d_1m, d_3m, self.d0, d_12m], start=6):
                val = self._chg(e, M_TRI, dd, self.d1)
                cell = V(ws.cell(row=r, column=k), val, fmt="+0.00;-0.00",
                         bold=(k == 8), color=C_HARD)
                if val is not None:
                    cell.fill = PatternFill("solid", fgColor=C_GOOD if val >= 0 else C_BAD)
            V(ws.cell(row=r, column=10), self._chg(e, M_STW, d_1m, self.d1), fmt="+0;-0")
            V(ws.cell(row=r, column=11), at.get("spread_bp"), fmt="+0;-0")
            V(ws.cell(row=r, column=12), at.get("ust_bp"), fmt="+0;-0")
            flag = at.get("flag") or ""
            c = V(ws.cell(row=r, column=13), flag, fmt="General", color=C_NOTE)
            c.alignment = Alignment(horizontal="left")
            F(c, italic=True, color=C_NOTE)
            r += 1
        ws.freeze_panes = "B5"
        r += 1
        ws.cell(row=r, column=1, value=(
            "Each return column header carries the date it is measured from, so there is never a "
            "question of which window a number belongs to. Returns are from the total-return "
            "sub-indices; spread and Treasury moves are chain-linked across basis breaks."))
        F(ws.cell(row=r, column=1), italic=True, color=C_NOTE)

    # ---------- month-end sampling ----------
    def month_ends(self, n: int = 60) -> List[str]:
        me: Dict[str, str] = {}
        for d in self.dates:
            me[d[:7]] = d
        picks = [me[k] for k in sorted(me)]
        if picks and picks[-1] != self.d1:
            picks.append(self.d1)
        return picks[-n:]

    def _chg(self, e: str, metric: str, d_from: Optional[str], d_to: str) -> Optional[float]:
        """Chain-linked change, so a basis break is not reported as a market move."""
        chained, _naive, _x = chain_change(self.p, e, metric, d_from, d_to, self.breaks)
        return chained

    def _back(self, months: int) -> Optional[str]:
        target = dt(self.d1) - timedelta(days=int(round(months * 30.44)))
        prior = [d for d in self.dates if dt(d) <= target]
        return prior[-1] if prior else None

    # ---------- the three time-series tabs ----------
    def _metric_tab(self, sheet: str, metric: str, title: str, unit: str,
                    level_fmt: str, chg_fmt: str):
        """Sub-indices first and summary columns on the LEFT, so performance
        reads without scrolling: level, 1m, 3m, YTD, 12m, then the history."""
        ws = self.wb.create_sheet(sheet)
        ws.sheet_view.showGridLines = False
        ws["A1"] = f"{title} — {self.d1}"
        F(ws["A1"], size=14, bold=True, color=C_HDR)
        ws["A2"] = (f"Sub-indices at the top: the index, then the five regions, then the rating "
                    f"buckets, then countries. The block on the left is how each one is doing "
                    f"({unit}); the columns to the right are month-end history.")
        F(ws["A2"], italic=True, color=C_NOTE)

        picks = self.month_ends()
        d_1m, d_3m, d_12m = self._back(1), self._back(3), self._back(12)
        left = [("Sub-index", 30), (f"Latest ({unit})", 14), ("1 month", 11),
                ("3 months", 11), (f"YTD {self.year}", 12), ("12 months", 11)]
        scn = metric == M_TRI
        if scn:
            left += [(f"{n} 12m %\n(Forecast tab)", 12) for n, _s in SCENARIOS]
        hr = 4
        for i, (h, w) in enumerate(left, start=1):
            H(ws.cell(row=hr, column=i), h)
            ws.column_dimensions[get_column_letter(i)].width = w
        for j, d in enumerate(picks, start=len(left) + 1):
            H(ws.cell(row=hr, column=j), d[:7] if d != self.d1 else d)
            ws.column_dimensions[get_column_letter(j)].width = 9
        ncols = len(left) + len(picks)

        r = hr + 1
        groups = [("INDEX", [INDEX]), ("REGIONS", REGIONS), ("RATING BUCKETS", RATINGS),
                  ("COUNTRIES", self.countries)]
        for gname, ents in groups:
            ents = [e for e in ents
                    if self.p.get(self.d1, e, metric) is not None
                    and (e == INDEX or (self.p.get(self.d1, e, M_WGT) or 0) > 0)]
            if not ents:
                continue
            r = band(ws, r, ncols, gname)
            for e in ents:
                is_agg = e == INDEX or e in REGIONS or e in RATINGS
                L(ws.cell(row=r, column=1), e, bold=is_agg,
                  fill=C_BAND if e == INDEX else None)
                V(ws.cell(row=r, column=2), self.p.get(self.d1, e, metric),
                  fmt=level_fmt, color=C_HARD, bold=is_agg)
                for k, dd in enumerate([d_1m, d_3m, self.d0, d_12m], start=3):
                    val = (self._chg(e, metric, dd, self.d1)
                           if metric == M_TRI or self.p.spread_meaningful(e) else None)
                    cell = V(ws.cell(row=r, column=k), val, fmt=chg_fmt, bold=is_agg)
                    if val is not None and abs(val) > 1e-9:
                        good = (val > 0) if metric == M_TRI else (val < 0)
                        cell.fill = PatternFill("solid", fgColor=C_GOOD if good else C_BAD)
                if scn:
                    y_ = self.p.get(self.d1, e, M_YLD)
                    s_ = self.p.get(self.d1, e, M_STW)
                    D_, _src = self.p.duration(self.d1, e)
                    IRD_ = self.p.ir_duration(self.d1, e)
                    for k, (n_, _s) in enumerate(SCENARIOS):
                        cell = ws.cell(row=r, column=7 + k)
                        if y_ is None or s_ is None or y_ > DISTRESSED_YIELD or not self.spread:
                            V(cell, None, fmt="0.00")
                            continue
                        col = SCN_COL[n_]
                        # own yield, own durations; the index's spread change scaled by
                        # this sub-index's spread relative to the index (proportional beta)
                        f_ = (f"={y_:.4f}-{IRD_:.4f}*Forecast!${col}${SCN_ROW_UST}/100"
                              f"-{D_:.4f}*(Forecast!${col}${SCN_ROW_CHG}*{s_:.2f}/{self.spread:.2f})/100")
                        V(cell, None, fmt="0.00", bold=is_agg, color=C_HARD)
                        cell.value = f_
                for j, d in enumerate(picks, start=len(left) + 1):
                    V(ws.cell(row=r, column=j), self.p.get(d, e, metric),
                      fmt=level_fmt, color=C_HARD)
                r += 1
            r += 1
        ws.freeze_panes = f"{get_column_letter(len(left) + 1)}{hr + 1}"

        note = ("Green is a gain, red a loss." if metric == M_TRI
                else "Green is tightening, red is widening.")
        ws.cell(row=r, column=1, value=note + "  Levels in the history block are month ends.")
        F(ws.cell(row=r, column=1), italic=True, color=C_NOTE)
        if scn:
            r += 1
            ws.cell(row=r, column=1, value=(
                "Scenario columns are live formulas on the Forecast tab's bear/base/bull inputs: each "
                "sub-index uses its own yield and durations, and takes the index spread change scaled by "
                "its own spread relative to the index (a 100bp index widening is 55bp for a 96bp IG bucket "
                "and 190bp for a 330bp B bucket). Blank where the yield is above "
                f"{DISTRESSED_YIELD:.0f}% — that is a recovery bet, not carry."))
            F(ws.cell(row=r, column=1), italic=True, color=C_NOTE)
            from openpyxl.formatting.rule import CellIsRule
            rng = f"G5:I{r}"
            ws.conditional_formatting.add(rng, CellIsRule(operator="greaterThanOrEqual", formula=[str(CASH_RATE)],
                                                          fill=PatternFill("solid", fgColor=C_OK)))
            ws.conditional_formatting.add(rng, CellIsRule(operator="lessThan", formula=[str(CASH_RATE)],
                                                          fill=PatternFill("solid", fgColor=C_NOTOK)))

    def spreads_tab(self):
        self._metric_tab("Spreads", M_STW, "Spreads — STW over Treasury", "bp", "0", "+0;-0")

    def yields_tab(self):
        self._metric_tab("Yields", M_YLD, "Yields to maturity", "%", "0.00", "+0;-0")

    def total_return_tab(self):
        self._metric_tab("Total_Return", M_TRI, "Total return index", "level", "0.00", "+0.00;-0.00")

    def _breaks_by_entity_block(self, ws, r: int) -> int:
        r = band(ws, r, 5, "Break days by sub-index — each series is tested on its own")
        for i, h in enumerate(["Sub-index", "Break days", "Last break", "Step that day (bp)",
                               "Restated share of level"], start=1):
            H(ws.cell(row=r, column=i), h)
        r += 1
        for e in [INDEX] + REGIONS + RATINGS + self.countries:
            b = self.p.breaks_for(e)
            if not b:
                continue
            last = sorted(b)[-1]
            L(ws.cell(row=r, column=1), e, bold=e == INDEX or e in REGIONS or e in RATINGS)
            V(ws.cell(row=r, column=2), len(b), fmt="0")
            V(ws.cell(row=r, column=3), last, fmt="@")
            V(ws.cell(row=r, column=4), b[last]["jump_bp"], fmt="+0;-0", color=C_HARD)
            sh = break_share(self.p, e)
            V(ws.cell(row=r, column=5), sh, fmt="0%", color=C_NOTOKTXT if sh > 0.5 else C_TXT)
            r += 1
        r += 1
        ws.cell(row=r, column=1, value=(
            "Every chain-linked spread change in this workbook skips that sub-index's own break days. "
            "The Middle East steps at month-ends as Lebanon is rebalanced; Latin America stepped when "
            "Venezuela was re-admitted in April 2024 and again when defaulted names left the spread on "
            "04-Sep-2026. None of those days is a market move."))
        F(ws.cell(row=r, column=1), italic=True, color=C_NOTE)
        return r

    # ---------- forecast ----------
    def forecast(self):
        ws = self.wb.create_sheet("Forecast")
        ws.sheet_view.showGridLines = False
        ws["A1"] = "12-month forecast"
        F(ws["A1"], size=14, bold=True, color=C_HDR)
        ws["A2"] = "Type your 12-month views into the shaded cells. Everything else calculates."
        F(ws["A2"], italic=True, color=C_NOTE)
        ws.column_dimensions["A"].width = 38
        for c in "BCDEFGHI":
            ws.column_dimensions[c].width = 14

        daily_s, direct_s, nobs = sigma_12m(self.p)
        r = 4
        r = band(ws, r, 3, "Current state")
        for lbl, v, fmt in [("Spread today (bp)", self.spread, "0"),
                            ("Index yield (%)", self.yld, "0.00"),
                            ("Spread duration", self.D, "0.00"),
                            ("IR duration", self.IRD, "0.00"),
                            ("Cash hurdle (%)", CASH_RATE, "0.00")]:
            L(ws.cell(row=r, column=1), lbl, bold=False)
            V(ws.cell(row=r, column=2), v, fmt=fmt, color=C_HARD, bold=True,
              fill=C_IN if "Cash" in lbl else None)
            r += 1
        S_REF, Y_REF, SD_REF, ID_REF, CASH_REF = "B5", "B6", "B7", "B8", "B9"
        r += 1

        r = band(ws, r, 4, "Your 12-month views")
        H(ws.cell(row=r, column=2), "your view")
        H(ws.cell(row=r, column=3), "beta")
        H(ws.cell(row=r, column=4), "bp of spread")
        r += 1
        first = r
        for lbl, key, fmt in [("UST 10y change (bp)", "ust10y_bp", "+0;-0"),
                              ("DXY change (%)", "dxy_pct", "+0.0;-0.0"),
                              ("VIX change (points)", "vix_pts", "+0.0;-0.0"),
                              ("EM rating drift (notches, + = upgrades)", "rating_notches", "+0.00;-0.00")]:
            L(ws.cell(row=r, column=1), lbl, bold=False)
            V(ws.cell(row=r, column=2), 0.0, fmt=fmt, fill=C_IN, bold=True)
            V(ws.cell(row=r, column=3), BETAS[key], fmt="0.00", color=C_NOTE)
            c = ws.cell(row=r, column=4, value=f"=B{r}*C{r}")
            V(c, None, fmt="+0.0;-0.0")
            c.value = f"=B{r}*C{r}"
            r += 1
        last = r - 1
        UST_REF = f"B{first}"
        L(ws.cell(row=r, column=1), "Total spread impact (bp)", bold=True)
        c = ws.cell(row=r, column=4, value=f"=SUM(D{first}:D{last})")
        V(c, None, fmt="+0.0;-0.0", bold=True); c.value = f"=SUM(D{first}:D{last})"
        DELTA = f"D{r}"
        r += 2

        r = band(ws, r, 4, "Forecast")
        for lbl, formula, fmt, bold in [
            ("Forecast spread (bp)", f"={S_REF}+{DELTA}", "0.0", True),
            ("Change from today (bp)", f"={DELTA}", "+0.0;-0.0", False),
            ("12m total return (%)", f"={Y_REF}-{ID_REF}*{UST_REF}/100-{SD_REF}*{DELTA}/100", "0.00", True),
            ("Excess over cash (pp)", f"={Y_REF}-{ID_REF}*{UST_REF}/100-{SD_REF}*{DELTA}/100-{CASH_REF}", "+0.00;-0.00", True),
            ("Spread at which return = cash (bp)", f"={S_REF}+({Y_REF}-{ID_REF}*{UST_REF}/100-{CASH_REF})*100/{SD_REF}", "0", False),
        ]:
            L(ws.cell(row=r, column=1), lbl, bold=False)
            c = ws.cell(row=r, column=2, value=formula)
            V(c, None, fmt=fmt, bold=bold, color=C_HARD); c.value = formula
            r += 1
        r += 1

        r = band(ws, r, 6, "Uncertainty around the central case")
        for lbl, v, note in [
            ("12m sigma, daily scaled by root-252 (bp)", daily_s,
             "overstates when spreads mean-revert"),
            ("12m sigma, actual 252-day changes (bp)", direct_s,
             f"from {nobs} overlapping windows — the honest one"),
        ]:
            L(ws.cell(row=r, column=1), lbl, bold=False)
            V(ws.cell(row=r, column=2), v, fmt="0", color=C_HARD)
            ws.cell(row=r, column=3, value=note)
            F(ws.cell(row=r, column=3), italic=True, color=C_NOTE)
            r += 1
        SIG = f"B{r - 1}"
        L(ws.cell(row=r, column=1), "Vol multiplier", bold=False)
        V(ws.cell(row=r, column=2), 1.0, fmt="0.00", fill=C_IN, bold=True)
        VM = f"B{r}"
        r += 2
        H(ws.cell(row=r, column=1), "Percentile")
        for i, p in enumerate(["p5", "p25", "p50", "p75", "p95"]):
            H(ws.cell(row=r, column=2 + i), p)
        r += 1
        L(ws.cell(row=r, column=1), "Spread (bp)", bold=False)
        for i, z in enumerate([-1.645, -0.674, 0.0, 0.674, 1.645]):
            f_ = f"={S_REF}+{DELTA}+{z}*{SIG}*{VM}"
            c = ws.cell(row=r, column=2 + i, value=f_)
            V(c, None, fmt="0", bold=(z == 0)); c.value = f_
        r += 1
        L(ws.cell(row=r, column=1), "Total return (%)", bold=False)
        for i, z in enumerate([-1.645, -0.674, 0.0, 0.674, 1.645]):
            f_ = (f"={Y_REF}-{ID_REF}*{UST_REF}/100-{SD_REF}*({DELTA}+{z}*{SIG}*{VM})/100")
            c = ws.cell(row=r, column=2 + i, value=f_)
            V(c, None, fmt="0.00", bold=(z == 0)); c.value = f_
        r += 2
        for line in [
            "Betas are practitioner priors, not regression estimates. The rates beta is",
            "unconditional: it cannot tell a rate rise driven by growth (spreads tighten) from",
            "one driven by an inflation shock (spreads widen). If your rate view depends on which,",
            "the forecast should be built regime by regime rather than off one coefficient.",
            "Bands are Gaussian and ignore fat tails; the vol multiplier is the crude lever.",
        ]:
            ws.cell(row=r, column=1, value=line)
            F(ws.cell(row=r, column=1), italic=True, color=C_NOTE)
            r += 1

        self._forecast_scenarios(ws, first, last)

    def _forecast_scenarios(self, ws, first: int, last: int) -> None:
        """Bear / base / bull side by side, to the right of the single-case plug. Live
        formulas: the betas are the same cells the central case uses, the spread can be
        set directly or left to the drivers, and the probability weights give an expected
        value. The Total_Return tab's scenario columns point at rows 6 and 13 here."""
        for c in "HIJKL":
            ws.column_dimensions[c].width = 15
        ws.column_dimensions["H"].width = 34
        cols = [SCN_COL[n] for n, _s in SCENARIOS]
        # band H4:L4
        L(ws["H4"], "12-month scenarios — bear / base / bull (shaded cells are yours)", fill=C_BAND)
        for c in "IJKL":
            ws[f"{c}4"].fill = PatternFill("solid", fgColor=C_BAND); ws[f"{c}4"].border = BRD
        H(ws["H5"], "")
        for (name, _s), c in zip(SCENARIOS, cols):
            H(ws[f"{c}5"], name)
        H(ws["L5"], "Prob-weighted")
        beta_rows = list(range(first, last + 1))      # UST, DXY, VIX, rating — same order
        drivers = [("UST 10y change (bp)", "ust", "+0;-0"), ("DXY change (%)", "dxy", "+0.0;-0.0"),
                   ("VIX change (points)", "vix", "+0.0;-0.0"),
                   ("EM rating drift (notches)", "rating", "+0.00;-0.00")]
        for k, (lbl, key, fmt) in enumerate(drivers):
            rr = 6 + k
            L(ws[f"H{rr}"], lbl, bold=False)
            for (name, s), c in zip(SCENARIOS, cols):
                V(ws[f"{c}{rr}"], s[key], fmt=fmt, fill=C_IN, bold=True)
            f_ = f"=SUMPRODUCT(I{rr}:K{rr},$I$16:$K$16)/SUM($I$16:$K$16)"
            V(ws[f"L{rr}"], None, fmt=fmt, color=C_NOTE); ws[f"L{rr}"].value = f_
        L(ws["H10"], "Spread from the drivers (bp)", bold=False)
        for c in cols:
            f_ = "=$B$5+" + "+".join(f"{c}{6 + k}*$C${beta_rows[k]}" for k in range(4))
            V(ws[f"{c}10"], None, fmt="0", color=C_NOTE); ws[f"{c}10"].value = f_
        L(ws["H11"], "Spread you want to argue (bp; blank = drivers)", bold=False)
        for (name, s), c in zip(SCENARIOS, cols):
            V(ws[f"{c}11"], s["spread"], fmt="0", fill=C_IN, bold=True)
        L(ws["H12"], "Spread used (bp)", bold=True)
        for c in cols:
            f_ = f'=IF({c}11="",{c}10,{c}11)'
            V(ws[f"{c}12"], None, fmt="0", bold=True, color=C_HARD); ws[f"{c}12"].value = f_
        f_ = "=SUMPRODUCT(I12:K12,$I$16:$K$16)/SUM($I$16:$K$16)"
        V(ws["L12"], None, fmt="0", bold=True, color=C_HARD); ws["L12"].value = f_
        L(ws["H13"], "Change vs today (bp)", bold=False)
        for c in cols:
            f_ = f"={c}12-$B$5"
            V(ws[f"{c}13"], None, fmt="+0;-0"); ws[f"{c}13"].value = f_
        L(ws["H14"], "12m total return (%)", bold=True)
        for c in cols:
            f_ = f"=$B$6-$B$8*{c}6/100-$B$7*{c}13/100"
            V(ws[f"{c}14"], None, fmt="0.00", bold=True, color=C_HARD); ws[f"{c}14"].value = f_
        f_ = "=SUMPRODUCT(I14:K14,$I$16:$K$16)/SUM($I$16:$K$16)"
        V(ws["L14"], None, fmt="0.00", bold=True, color=C_HARD); ws["L14"].value = f_
        L(ws["H15"], "Excess over cash (pp)", bold=False)
        for c in cols + ["L"]:
            f_ = f"={c}14-$B$9"
            V(ws[f"{c}15"], None, fmt="+0.00;-0.00"); ws[f"{c}15"].value = f_
        L(ws["H16"], "Probability (%)", bold=False)
        for (name, s), c in zip(SCENARIOS, cols):
            V(ws[f"{c}16"], s["prob"], fmt="0", fill=C_IN, bold=True)
        V(ws["L16"], None, fmt="0", color=C_NOTE); ws["L16"].value = "=SUM(I16:K16)"
        L(ws["H17"], "Spread at which return = cash (bp)", bold=False)
        for c in cols:
            f_ = f"=$B$5+($B$6-$B$8*{c}6/100-$B$9)*100/$B$7"
            V(ws[f"{c}17"], None, fmt="0", color=C_HARD); ws[f"{c}17"].value = f_
        # Excel-native conditional colour so the cells recolour as the inputs change.
        from openpyxl.formatting.rule import CellIsRule
        ws.conditional_formatting.add("I14:L14", CellIsRule(operator="greaterThanOrEqual", formula=["$B$9"],
                                                             fill=PatternFill("solid", fgColor=C_OK)))
        ws.conditional_formatting.add("I14:L14", CellIsRule(operator="lessThan", formula=["$B$9"],
                                                             fill=PatternFill("solid", fgColor=C_NOTOK)))
        from openpyxl.chart import BarChart
        ch = BarChart()
        ch.type = "col"
        ch.title = "12m total return by scenario (%)"
        ch.add_data(Reference(ws, min_col=9, max_col=11, min_row=14, max_row=14), from_rows=True, titles_from_data=False)
        ch.set_categories(Reference(ws, min_col=9, max_col=11, min_row=5, max_row=5))
        ch.legend = None
        ch.height, ch.width = 6.5, 11
        ws.add_chart(ch, "H19")
        ws["H33"] = ("The base case is the spread you defend, the bear case is the official number, the bull "
                     "case is what a Treasury rally does. Change any shaded cell; the Total_Return tab's "
                     "scenario columns and the probability-weighted column follow.")
        F(ws["H33"], italic=True, color=C_NOTE)

    # ---------- sensitivity ----------
    def sensitivity(self):
        ws = self.wb.create_sheet("Sensitivity")
        ws.sheet_view.showGridLines = False
        ws["A1"] = "Sensitivity to the US 10-year"
        F(ws["A1"], size=14, bold=True, color=C_HDR)
        ws["A2"] = (f"Spot spread {self.spread:.0f}bp, yield {self.yld:.2f}%, spread duration "
                    f"{self.D:.2f}, IR duration {self.IRD:.2f}, cash {CASH_RATE:.2f}%.")
        F(ws["A2"], italic=True, color=C_NOTE)
        ws.column_dimensions["A"].width = 22
        for i in range(2, 12):
            ws.column_dimensions[get_column_letter(i)].width = 13

        r = 4
        ws.cell(row=3, column=1, value=(
            f"US 10y {self.ust10y:.2f}%  ·  index Treasury anchor {self.ust_now:.2f}%  "
            f"(term premium {self.term_premium * 100:+.0f}bp for {10.2:.1f}y average life)  ·  "
            f"curve beta {CURVE_BETA:.2f}"))
        F(ws.cell(row=3, column=1), italic=True, color=C_NOTE)
        r = band(ws, r, 7, f"A. Spread responds to rates through beta ({BETAS['ust10y_bp']:.2f})")
        for i, h in enumerate(["US 10y level", "10y change (bp)", "index anchor (%)",
                               "implied spread (bp)", "spread change (bp)",
                               "12m total return %", "vs cash (pp)"], start=1):
            H(ws.cell(row=r, column=i), h)
        r += 1
        for lvl in self.ust_grid:
            u = (lvl - self.ust10y) * 100.0            # move in the 10y
            anchor = lvl + self.term_premium + (CURVE_BETA - 1.0) * u / 100.0
            du = (anchor - self.ust_now) * 100.0       # move the index actually feels
            sp = self.spread + BETAS["ust10y_bp"] * u
            tr = total_return(self.yld, self.IRD, self.D, du, self.spread, sp)
            L(ws.cell(row=r, column=1), f"{lvl:.2f}%",
              bold=abs(u) < 1e-6, fill=C_BAND if abs(u) < 12.5 else None)
            V(ws.cell(row=r, column=2), u, fmt="+0;-0")
            V(ws.cell(row=r, column=3), anchor, fmt="0.00", color=C_HARD)
            V(ws.cell(row=r, column=4), sp, fmt="0", color=C_HARD)
            V(ws.cell(row=r, column=5), sp - self.spread, fmt="+0;-0")
            V(ws.cell(row=r, column=6), tr, fmt="0.00", bold=True,
              fill=C_GOOD if tr >= CASH_RATE else C_BAD)
            V(ws.cell(row=r, column=7), tr - CASH_RATE, fmt="+0.00;-0.00")
            r += 1
        r += 1

        r = band(ws, r, len(SENSITIVITY_UST) + 1,
                 "B. Spread and rates set independently — 12m total return (%)")
        H(ws.cell(row=r, column=1), "spread \\ UST")
        for j, u in enumerate(SENSITIVITY_UST, start=2):
            H(ws.cell(row=r, column=j), f"{u:+d}bp")
        r += 1
        for sp in SENSITIVITY_SPREADS:
            L(ws.cell(row=r, column=1), f"{sp}bp", bold=False)
            for j, u in enumerate(SENSITIVITY_UST, start=2):
                tr = total_return(self.yld, self.IRD, self.D, u, self.spread, sp)
                V(ws.cell(row=r, column=j), tr, fmt="0.00",
                  fill=C_GOOD if tr >= CASH_RATE else C_BAD)
            r += 1
        r += 1
        ws.cell(row=r, column=1, value=(
            "The index anchor column is the Treasury the index is actually exposed to: the 10y "
            "plus the term premium for its 10.2y average life. At a curve beta of 1.00 the "
            "premium cancels out of every change, so it is presentational; set the beta away "
            "from 1.00 to model a flattening or steepening and it starts to bite."))
        F(ws.cell(row=r, column=1), italic=True, color=C_NOTE)
        r += 2
        L(ws.cell(row=r, column=1), "break-even spread", bold=True)
        for j, u in enumerate(SENSITIVITY_UST, start=2):
            be = self.spread + (self.yld - self.IRD * (u / 100.0) - CASH_RATE) * 100.0 / self.D
            V(ws.cell(row=r, column=j), be, fmt="0", bold=True, color=C_HARD)
        r += 2
        ws.cell(row=r, column=1, value=("Green clears the cash hurdle, red does not. Grid B holds the "
                                        "two variables independent, which is conservative: in a "
                                        "genuine risk-off the Treasury rally offsets part of the "
                                        "spread widening, so the bottom-left corner overstates the "
                                        "pain. The break-even row is the number that actually "
                                        "decides the call."))
        F(ws.cell(row=r, column=1), italic=True, color=C_NOTE)

    # ---------- scenario grid ----------
    def scenario_tab(self):
        ws = self.wb.create_sheet("Scenario_Grid")
        ws.sheet_view.showGridLines = False
        ws["A1"] = "Scenario grid — 12-month total return by US 10y level and index spread level"
        F(ws["A1"], size=14, bold=True, color=C_HDR)
        ws["A2"] = (f"Spot: spread {self.spread:.0f}bp, yield {self.yld:.2f}%, US 10y "
                    f"{self.ust10y:.2f}%, spread duration {self.D:.2f}, IR duration {self.IRD:.2f}, "
                    f"cash {CASH_RATE:.2f}%. Return = yield − IR duration × Δ10y − spread duration "
                    f"× Δspread, spread and rates set independently.")
        F(ws["A2"], italic=True, color=C_NOTE)
        ws["A3"] = (f"Blue: beats cash by {GRID_MARGIN_PP:.1f}pp or more (ok). Grey: within "
                    f"{GRID_MARGIN_PP:.1f}pp of cash either side (marginal). Red: below cash by "
                    f"more than {GRID_MARGIN_PP:.1f}pp (not ok). Bold outline marks the spot "
                    f"10y column; the first row is the spot spread.")
        F(ws["A3"], italic=True, color=C_NOTE)
        ws.column_dimensions["A"].width = 18
        levels, rows = scenario_grid(self)
        for i in range(2, len(levels) + 3):
            ws.column_dimensions[get_column_letter(i)].width = 11
        spot_col = None
        for j, lvl in enumerate(levels, start=2):
            if abs(lvl - self.ust10y) <= 0.005:
                spot_col = j
        thick = Side(style="medium", color=C_HDR)

        def lvl_label(lvl):
            return f"{lvl:.2f}%" + (" spot" if abs(lvl - self.ust10y) <= 0.005 else "")

        def grid(r, title, values_fn, fmt):
            r = band(ws, r, len(levels) + 1, title)
            H(ws.cell(row=r, column=1), "spread \\ US 10y")
            for j, lvl in enumerate(levels, start=2):
                H(ws.cell(row=r, column=j), lvl_label(lvl))
            r += 1
            for lab, s_, vals in rows:
                L(ws.cell(row=r, column=1), lab, bold=lab.startswith("Spot"),
                  fill=C_BAND if lab.startswith("Spot") else None)
                for j, tr in enumerate(vals, start=2):
                    v = grid_verdict(tr)
                    fill = {"ok": C_OK, "marginal": C_MARG, "bad": C_NOTOK}[v]
                    col = {"ok": C_OKTXT, "marginal": C_TXT, "bad": C_NOTOKTXT}[v]
                    c = V(ws.cell(row=r, column=j), values_fn(tr), fmt=fmt, fill=fill,
                          color=col, bold=(j == spot_col))
                    if j == spot_col:
                        c.border = Border(left=thick, right=thick,
                                          top=Side(style="thin", color=C_BRD),
                                          bottom=Side(style="thin", color=C_BRD))
                r += 1
            return r + 1

        r = grid(5, "A. 12-month total return (%)", lambda t: t, "0.00")
        r = grid(r, "B. Excess over cash (pp)", lambda t: t - CASH_RATE, "+0.00;-0.00")

        r = band(ws, r, len(levels) + 1, "C. Break-even spread — the level at which the return equals cash")
        L(ws.cell(row=r, column=1), "US 10y level")
        for j, lvl in enumerate(levels, start=2):
            H(ws.cell(row=r, column=j), lvl_label(lvl))
        r += 1
        L(ws.cell(row=r, column=1), "break-even (bp)")
        for j, lvl in enumerate(levels, start=2):
            u = (lvl - self.ust10y) * 100.0
            be = self.spread + (self.yld - self.IRD * (u / 100.0) - CASH_RATE) * 100.0 / self.D
            V(ws.cell(row=r, column=j), be, fmt="0", bold=True, color=C_HARD)
        r += 1
        L(ws.cell(row=r, column=1), "room vs spot (bp)")
        for j, lvl in enumerate(levels, start=2):
            u = (lvl - self.ust10y) * 100.0
            be = self.spread + (self.yld - self.IRD * (u / 100.0) - CASH_RATE) * 100.0 / self.D
            V(ws.cell(row=r, column=j), be - self.spread, fmt="+0;-0")
        r += 2
        for line in [
            "How to read it: pick the 10y level you believe in, read down to the spread you expect, "
            "and the cell is the 12-month total return. Row C is the spread the index can widen to "
            "before the call stops beating cash at that 10y level; 'room vs spot' is how many basis "
            "points of widening that leaves.",
            "The grid holds spread and rates independent, which is conservative in the bottom-right: "
            "in a real risk-off the Treasury rally offsets part of the widening. The Sensitivity tab "
            "links them through the rates beta instead.",
            "Convexity is ignored; below roughly 50bp of yield change it is worth under 0.05%.",
        ]:
            ws.cell(row=r, column=1, value=line)
            F(ws.cell(row=r, column=1), italic=True, color=C_NOTE)
            r += 1

    # ---------- shared data for the new tabs and slides ----------
    def segments(self) -> List[str]:
        """Index, regions and the rating buckets that exist today (no countries)."""
        return [e for e in [INDEX] + REGIONS + RATINGS
                if self.p.get(self.d1, e, M_YLD) is not None
                and (e == INDEX or (self.p.get(self.d1, e, M_WGT) or 0) > 0)]

    @staticmethod
    def short(e: str) -> str:
        return (e.replace(" Region", "").replace("Credit ", "").replace(" only", "")
                .replace("EMBIG Div", "EMBIGD"))

    def cushion_rows(self) -> List[Dict[str, Any]]:
        """How many basis points of widening the next twelve months of carry can absorb
        before the segment stops beating cash. (yield - cash) / spread duration."""
        out = []
        for e in self.segments():
            y = self.p.get(self.d1, e, M_YLD) or 0.0
            s = self.p.get(self.d1, e, M_STW) or 0.0
            D, _src = self.p.duration(self.d1, e)
            IRD = self.p.ir_duration(self.d1, e)
            cush = (y - CASH_RATE) * 100.0 / D if D else None
            cush50 = (y - IRD * 0.5 - CASH_RATE) * 100.0 / D if D else None
            out.append(dict(e=e, name=self.short(e), wgt=self.p.get(self.d1, e, M_WGT),
                            yld=y, spread=s, D=D, IRD=IRD, cushion=cush, cushion_up50=cush50,
                            cushion_pct=(cush / s * 100.0) if (cush is not None and s) else None,
                            distressed=y > DISTRESSED_YIELD))
        return out

    def spread_stats(self) -> List[Dict[str, Any]]:
        """Today's spread against its own restated history: 1y and full-sample range,
        median, and the percentile today sits at."""
        d1y = self._back(12)
        out = []
        for e in self.segments():
            cs = chained_series(self.p, e, M_STW, self.breaks)
            if len(cs) < 60:
                continue
            allv = list(cs.values())
            y1 = [v for d, v in cs.items() if d1y and d >= d1y]
            now = self.p.get(self.d1, e, M_STW)
            share = break_share(self.p, e, self.breaks)
            out.append(dict(e=e, name=self.short(e), now=now, n_days=len(allv), share=share,
                            indicative=share > 0.5,
                            since=min(cs), lo=min(allv), p25=pctl(allv, 0.25),
                            med=pctl(allv, 0.5), p75=pctl(allv, 0.75), hi=max(allv),
                            lo1y=min(y1) if y1 else None, hi1y=max(y1) if y1 else None,
                            rank=pct_rank(allv, now) if now is not None else None))
        return out

    def return_stats(self) -> Dict[str, Any]:
        rr = rolling_12m_returns(self.p, INDEX)
        vals = [v for _d, v in rr]
        if not vals:
            return {}
        lo, hi = math.floor(min(vals) / 2.0) * 2, math.ceil(max(vals) / 2.0) * 2
        buckets = []
        x = lo
        while x < hi:
            buckets.append((x, sum(1 for v in vals if x <= v < x + 2)))
            x += 2
        return dict(n=len(vals), first=rr[0][0], last=rr[-1][0],
                    p10=pctl(vals, 0.10), p25=pctl(vals, 0.25), med=pctl(vals, 0.5),
                    p75=pctl(vals, 0.75), p90=pctl(vals, 0.90), lo=min(vals), hi=max(vals),
                    lo_start=min(rr, key=lambda t: t[1])[0], hi_start=max(rr, key=lambda t: t[1])[0],
                    beat_cash=100.0 * sum(1 for v in vals if v > CASH_RATE) / len(vals),
                    negative=100.0 * sum(1 for v in vals if v < 0) / len(vals),
                    buckets=buckets, drawdowns=worst_drawdowns(self.p, INDEX, 3),
                    latest=self._chg(INDEX, M_TRI, self._back(12), self.d1))

    def jpm_legs(self) -> List[Dict[str, Any]]:
        """JPM's own YTD decomposition from the latest snapshot, if one is within a week
        of the last date: total = (1 + excess over Treasuries) x (1 + Treasury return) - 1."""
        cands = [(abs((dt(k) - dt(self.d1)).days), k) for k in self.p.extras]
        if not cands or min(cands)[0] > 7:
            return []
        k = min(cands)[1]
        ex = self.p.extras[k]
        out = []
        for e in self.segments():
            g = lambda f: num(ex.get((e, f)))
            tot, sp, us = g("YTD Change (%)"), g("Spread Return YTD Change (%)"), g("UST Return YTD Change (%)")
            if None in (tot, sp, us):
                continue
            out.append(dict(e=e, name=self.short(e), total=tot, spread=sp, ust=us,
                            coupon=g("Coupon Return YTD Change (%)"),
                            price=g("Price Return YTD Change (%)"),
                            cross=tot - sp - us, asof=k))
        return out

    def scenario_values(self) -> List[Dict[str, Any]]:
        """The Forecast tab's bear/base/bull defaults, evaluated in Python for the deck.
        The workbook version is live; this one mirrors its defaults."""
        out = []
        for name, s in SCENARIOS:
            drv = (self.spread + BETAS["ust10y_bp"] * s["ust"] + BETAS["dxy_pct"] * s["dxy"]
                   + BETAS["vix_pts"] * s["vix"] + BETAS["rating_notches"] * s["rating"])
            sp = s["spread"] if s["spread"] is not None else drv
            tr = total_return(self.yld, self.IRD, self.D, s["ust"], self.spread, sp)
            out.append(dict(name=name, ust=s["ust"], spread=sp, drivers=drv, tr=tr,
                            excess=tr - CASH_RATE, prob=s["prob"]))
        return out

    # ---------- YTD attribution, JPM's legs, by segment ----------
    def attribution_tab(self):
        legs = self.jpm_legs()
        if not legs:
            return
        ws = self.wb.create_sheet("YTD_Attribution")
        ws.sheet_view.showGridLines = False
        ws["A1"] = f"What explained the year — JPM's published decomposition, as of {legs[0]['asof']}"
        F(ws["A1"], size=14, bold=True, color=C_HDR)
        ws["A2"] = ("JPM splits each sub-index's YTD total return into the excess return over "
                    "duration-matched Treasuries (spread move plus spread carry) and the Treasury "
                    "return (rate move plus Treasury carry). They compound: total = (1 + excess) x "
                    "(1 + Treasury) - 1, so the cross term is shown rather than hidden.")
        F(ws["A2"], italic=True, color=C_NOTE)
        ws.column_dimensions["A"].width = 22
        for c in "BCDEFGH":
            ws.column_dimensions[c].width = 15
        r = 4
        for i, h in enumerate(["Segment", "Weight %", "Total YTD %", "Excess over UST %",
                               "Treasury %", "Cross term %", "Coupon %", "Price %"], start=1):
            H(ws.cell(row=r, column=i), h)
        r += 1
        for L_ in legs:
            is_idx = L_["e"] == INDEX
            L(ws.cell(row=r, column=1), L_["name"], bold=is_idx, fill=C_BAND if is_idx else None)
            V(ws.cell(row=r, column=2), self.p.get(self.d1, L_["e"], M_WGT), fmt="0.0")
            V(ws.cell(row=r, column=3), L_["total"], fmt="+0.00;-0.00", bold=True,
              fill=C_GOOD if L_["total"] >= 0 else C_BAD)
            V(ws.cell(row=r, column=4), L_["spread"], fmt="+0.00;-0.00", color=C_HARD)
            V(ws.cell(row=r, column=5), L_["ust"], fmt="+0.00;-0.00", color=C_HARD)
            V(ws.cell(row=r, column=6), L_["cross"], fmt="+0.00;-0.00", color=C_NOTE)
            V(ws.cell(row=r, column=7), L_["coupon"], fmt="+0.00;-0.00")
            V(ws.cell(row=r, column=8), L_["price"], fmt="+0.00;-0.00")
            r += 1
        r += 1
        for line in [
            "Read across: a segment whose excess-over-Treasuries leg is large and positive while the "
            "Treasury leg is negative earned its return from spreads and carry against a rates headwind.",
            "This is JPM's arithmetic, not ours — the duration-based decomposition was removed because "
            "it produced nonsense for distressed names. For Credit C and single names the cross term is "
            "large because the legs are large; that is compounding, not an error.",
        ]:
            ws.cell(row=r, column=1, value=line)
            F(ws.cell(row=r, column=1), italic=True, color=C_NOTE)
            r += 1

    # ---------- cushion ----------
    def cushion_tab(self):
        rows = self.cushion_rows()
        ws = self.wb.create_sheet("Cushion")
        ws.sheet_view.showGridLines = False
        ws["A1"] = "Cushion — how much widening each segment can absorb before it loses to cash"
        F(ws["A1"], size=14, bold=True, color=C_HDR)
        ws["A2"] = (f"Cushion (bp) = (yield - cash {CASH_RATE:.2f}%) / spread duration x 100. It is the "
                    f"spread widening over the next twelve months that would bring the total return "
                    f"exactly to cash, rates unchanged. The second cushion assumes the 10y rises 50bp.")
        F(ws["A2"], italic=True, color=C_NOTE)
        ws.column_dimensions["A"].width = 22
        for c in "BCDEFGHI":
            ws.column_dimensions[c].width = 14
        r = 4
        for i, h in enumerate(["Segment", "Weight %", "Yield %", "Spread bp", "Spread dur",
                               "Cushion bp", "Cushion if 10y +50bp", "Cushion / spread %",
                               "12m return if nothing moves %"], start=1):
            H(ws.cell(row=r, column=i), h)
        r += 1
        for c_ in rows:
            is_idx = c_["e"] == INDEX
            L(ws.cell(row=r, column=1), c_["name"], bold=is_idx, fill=C_BAND if is_idx else None)
            V(ws.cell(row=r, column=2), c_["wgt"], fmt="0.0")
            V(ws.cell(row=r, column=3), c_["yld"], fmt="0.00")
            V(ws.cell(row=r, column=4), c_["spread"], fmt="0")
            V(ws.cell(row=r, column=5), c_["D"], fmt="0.00")
            for j, key in [(6, "cushion"), (7, "cushion_up50")]:
                v = c_[key]
                V(ws.cell(row=r, column=j), v, fmt="+0;-0", bold=(j == 6),
                  fill=None if v is None else (C_GOOD if v >= 50 else (C_BAD if v < 0 else None)),
                  color=C_NOTE if c_["distressed"] else C_TXT)
            V(ws.cell(row=r, column=8), c_["cushion_pct"], fmt="0")
            V(ws.cell(row=r, column=9), c_["yld"], fmt="0.00", color=C_NOTE if c_["distressed"] else C_TXT)
            r += 1
        r += 1
        for line in [
            "Blue: at least 50bp of room. Red: already below cash at today's yield. Grey text: yield above "
            f"{DISTRESSED_YIELD:.0f}% is a recovery bet, not a carry you will collect.",
            "The point for the committee: the argument for EM against cash is not the spread forecast, it is "
            "how far spreads can go wrong before the carry is gone. That is the cushion column.",
        ]:
            ws.cell(row=r, column=1, value=line)
            F(ws.cell(row=r, column=1), italic=True, color=C_NOTE)
            r += 1

    # ---------- spread vs own history ----------
    def spread_history_tab(self):
        st = self.spread_stats()
        if not st:
            return
        ws = self.wb.create_sheet("Spread_vs_History")
        ws.sheet_view.showGridLines = False
        ws["A1"] = "Spreads against their own history — restated on today's basis"
        F(ws["A1"], size=14, bold=True, color=C_HDR)
        ws["A2"] = ("History is rebuilt from today's published level backwards using chain-linked daily "
                    "moves, so the basis-break steps (including JPM's 04-Sep exclusion of defaulted names) "
                    "are removed proportionally. Without this, today would read as a record tight for the wrong "
                    "reason. It is an approximation; the last column says how much to trust each row.")
        F(ws["A2"], italic=True, color=C_NOTE)
        ws.column_dimensions["A"].width = 22
        for c in "BCDEFGHIJKL":
            ws.column_dimensions[c].width = 12
        r = 4
        for i, h in enumerate(["Segment", "Now bp", "1y low", "1y high", "Full low", "25th pct",
                               "Median", "75th pct", "Full high", "Percentile now", "History from",
                               "Break steps / level"], start=1):
            H(ws.cell(row=r, column=i), h)
        r += 1
        for s_ in st:
            is_idx = s_["e"] == INDEX
            L(ws.cell(row=r, column=1), s_["name"], bold=is_idx, fill=C_BAND if is_idx else None)
            V(ws.cell(row=r, column=2), s_["now"], fmt="0", bold=True, color=C_HARD)
            for j, key in enumerate(["lo1y", "hi1y", "lo", "p25", "med", "p75", "hi"], start=3):
                V(ws.cell(row=r, column=j), s_[key], fmt="0")
            rk = s_["rank"]
            V(ws.cell(row=r, column=10), rk, fmt="0", bold=True,
              fill=None if rk is None else (C_BAD if rk <= 10 else (C_GOOD if rk >= 50 else None)))
            V(ws.cell(row=r, column=11), s_["since"], fmt="@")
            V(ws.cell(row=r, column=12), s_["share"], fmt="0%", color=C_NOTE if not s_["indicative"] else C_NOTOKTXT,
              bold=s_["indicative"])
            r += 1
        r += 1
        for line in [
            "Percentile is the share of days in the restated history with a spread at or below today's. "
            "Red: inside the tightest tenth of history — little room left from valuation alone. "
            "Blue: at or above the median — spreads are not the constraint.",
            "This is the slide that answers 'aren't spreads already too tight?' with a number instead of an adjective.",
            "The last column is the size of the removed basis-break steps relative to today's level. Where it is "
            "above 50% (red) the segment's composition changed too much — Lebanon in the Middle East, Russia and "
            "Ukraine in Europe, the C bucket — and the restated history is indicative only.",
        ]:
            ws.cell(row=r, column=1, value=line)
            F(ws.cell(row=r, column=1), italic=True, color=C_NOTE)
            r += 1

    # ---------- what history says ----------
    def return_history_tab(self):
        rs = self.return_stats()
        if not rs:
            return
        ws = self.wb.create_sheet("History_Says")
        ws.sheet_view.showGridLines = False
        ws["A1"] = "What history says — every rolling 12-month total return in the sample"
        F(ws["A1"], size=14, bold=True, color=C_HDR)
        ws["A2"] = (f"{rs['n']} overlapping 12-month windows starting {rs['first']} to {rs['last']}, "
                    f"from the index's own total return series. Overlapping windows are not independent "
                    f"observations; read the percentiles as a description of the sample, not a probability.")
        F(ws["A2"], italic=True, color=C_NOTE)
        ws.column_dimensions["A"].width = 34
        for c in "BCDEFGH":
            ws.column_dimensions[c].width = 13
        r = 4
        r = band(ws, r, 3, "Distribution of 12-month total returns (%)")
        for lbl, key, fmt in [("Worst window", "lo", "+0.00;-0.00"), ("10th percentile", "p10", "+0.00;-0.00"),
                              ("25th percentile", "p25", "+0.00;-0.00"), ("Median", "med", "+0.00;-0.00"),
                              ("75th percentile", "p75", "+0.00;-0.00"), ("90th percentile", "p90", "+0.00;-0.00"),
                              ("Best window", "hi", "+0.00;-0.00"),
                              ("Share of windows beating cash (%)", "beat_cash", "0"),
                              ("Share of windows negative (%)", "negative", "0"),
                              ("Latest 12 months", "latest", "+0.00;-0.00")]:
            L(ws.cell(row=r, column=1), lbl, bold=key in ("med", "latest"))
            V(ws.cell(row=r, column=2), rs[key], fmt=fmt, bold=key in ("med", "latest"), color=C_HARD)
            if key == "lo":
                ws.cell(row=r, column=3, value=f"window starting {rs['lo_start']}")
            if key == "hi":
                ws.cell(row=r, column=3, value=f"window starting {rs['hi_start']}")
            F(ws.cell(row=r, column=3), italic=True, color=C_NOTE)
            r += 1
        r += 1
        r = band(ws, r, 3, "Histogram — count of windows by 2pp bucket")
        H(ws.cell(row=r, column=1), "Bucket"); H(ws.cell(row=r, column=2), "Windows")
        r += 1
        h0 = r
        for lo, n in rs["buckets"]:
            L(ws.cell(row=r, column=1), f"{lo:+.0f} to {lo + 2:+.0f}", bold=False)
            V(ws.cell(row=r, column=2), n, fmt="0")
            r += 1
        h1 = r - 1
        from openpyxl.chart import BarChart
        ch = BarChart()
        ch.type = "col"
        ch.title = "Rolling 12m total returns, count by bucket"
        ch.add_data(Reference(ws, min_col=2, min_row=h0 - 1, max_row=h1), titles_from_data=True)
        ch.set_categories(Reference(ws, min_col=1, min_row=h0, max_row=h1))
        ch.legend = None
        ch.height, ch.width = 7.5, 16
        ws.add_chart(ch, "E4")
        r += 1
        r = band(ws, r, 6, "Worst drawdowns in the total return index")
        for i, h in enumerate(["Peak", "Trough", "Depth %", "Days to trough", "Recovered", "Days to recover"], start=1):
            H(ws.cell(row=r, column=i), h)
        r += 1
        for dd in rs["drawdowns"]:
            L(ws.cell(row=r, column=1), dd["peak"], bold=False)
            V(ws.cell(row=r, column=2), dd["trough"], fmt="@")
            V(ws.cell(row=r, column=3), dd["depth"], fmt="0.00", fill=C_BAD)
            V(ws.cell(row=r, column=4), dd["days_down"], fmt="0")
            V(ws.cell(row=r, column=5), dd["recovered"] or "not yet", fmt="@")
            V(ws.cell(row=r, column=6), dd["days_back"], fmt="0")
            r += 1
        r += 1
        scn = self.scenario_values()
        r = band(ws, r, 6, "Where the scenarios sit (defaults from the Forecast tab)")
        for i, h in enumerate(["Scenario", "10y change bp", "Spread bp", "12m return %", "vs cash pp", "Percentile of history"], start=1):
            H(ws.cell(row=r, column=i), h)
        r += 1
        allv = [v for _d, v in rolling_12m_returns(self.p, INDEX)]
        for s_ in scn:
            L(ws.cell(row=r, column=1), s_["name"], bold=s_["name"] == "Base")
            V(ws.cell(row=r, column=2), s_["ust"], fmt="+0;-0")
            V(ws.cell(row=r, column=3), s_["spread"], fmt="0")
            V(ws.cell(row=r, column=4), s_["tr"], fmt="0.00", bold=True,
              fill=C_GOOD if s_["tr"] >= CASH_RATE else C_BAD)
            V(ws.cell(row=r, column=5), s_["excess"], fmt="+0.00;-0.00")
            V(ws.cell(row=r, column=6), pct_rank(allv, s_["tr"]), fmt="0")
            r += 1
        r += 1
        ws.cell(row=r, column=1, value=(
            "Use: a base case sitting near the historical median is easy to defend; one in the top decile "
            "needs a reason. The drawdown table is the answer to 'what if we are wrong' — depth and time to "
            "recover, from this index's own record."))
        F(ws.cell(row=r, column=1), italic=True, color=C_NOTE)

    # ---------- basis breaks ----------
    def breaks_tab(self):
        ws = self.wb.create_sheet("Basis_Breaks")
        ws.sheet_view.showGridLines = False
        ws["A1"] = "Basis breaks — days the spread series stepped without the market moving"
        F(ws["A1"], size=14, bold=True, color=C_HDR)
        ws["A2"] = ("Detected automatically: a spread jump with no matching move in the total "
                    "return index. Every period change in this workbook is chain-linked across "
                    "these days, so what you see is the market, not the methodology.")
        F(ws["A2"], italic=True, color=C_NOTE)
        ws.column_dimensions["A"].width = 16
        for c in "BCDEFG":
            ws.column_dimensions[c].width = 17

        hr = 4
        for i, h in enumerate(["Date", "Spread jump (bp)", "Actual return %",
                               "Return the jump implied %", "Gap"], start=1):
            H(ws.cell(row=hr, column=i), h)
        r = hr + 1
        for d in sorted(self.breaks):
            b = self.breaks[d]
            L(ws.cell(row=r, column=1), d, bold=False)
            V(ws.cell(row=r, column=2), b["jump_bp"], fmt="+0;-0", color=C_HARD, bold=True)
            V(ws.cell(row=r, column=3), b["actual_pct"], fmt="+0.00;-0.00")
            V(ws.cell(row=r, column=4), b["implied_pct"], fmt="+0.00;-0.00")
            V(ws.cell(row=r, column=5), b["actual_pct"] - b["implied_pct"], fmt="+0.00;-0.00",
              fill=C_BAD)
            r += 1
        if not self.breaks:
            ws.cell(row=r, column=1, value="None detected.")
            F(ws.cell(row=r, column=1), italic=True, color=C_NOTE)
            return
        r += 2

        latest = sorted(self.breaks)[-1]
        prior = [d for d in self.dates if d < latest][-1]
        r = band(ws, r, 5, f"Who moved on {latest} — ratio of new basis to old")
        for i, h in enumerate(["Sub-index", "before", "after", "change", "ratio"], start=1):
            H(ws.cell(row=r, column=i), h)
        r += 1
        rows = []
        for e in [INDEX] + REGIONS + RATINGS + self.countries:
            a, b = self.p.get(prior, e, M_STW), self.p.get(latest, e, M_STW)
            if a is None or b is None or a == 0:
                continue
            rows.append((b / a, e, a, b))
        rows.sort()
        for ratio, e, a, b in rows[:14] + [(None, "", None, None)] + rows[-6:]:
            if not e:
                r += 1
                continue
            L(ws.cell(row=r, column=1), e, bold=(e == INDEX or e in REGIONS or e in RATINGS))
            V(ws.cell(row=r, column=2), a, fmt="0", color=C_HARD)
            V(ws.cell(row=r, column=3), b, fmt="0", color=C_HARD)
            V(ws.cell(row=r, column=4), b - a, fmt="+0;-0")
            V(ws.cell(row=r, column=5), ratio, fmt="0.00",
              fill=C_BAD if ratio < 0.95 else None)
            r += 1
        r += 1
        for line in [
            "Read the ratio column. Performing credits sit at 1.00 — untouched. Defaulted and",
            "near-default paper is roughly halved. That is the signature of a change in how",
            "spread is computed for non-performing bonds, not of anything that happened in the",
            "market, and it is why the naive year-to-date spread move is not the market's move.",
        ]:
            ws.cell(row=r, column=1, value=line)
            F(ws.cell(row=r, column=1), italic=True, color=C_NOTE)
            r += 1
        r += 1
        self._breaks_by_entity_block(ws, r)

    # ---------- methodology ----------
    def methodology(self):
        ws = self.wb.create_sheet("Methodology")
        ws.sheet_view.showGridLines = False
        ws.column_dimensions["A"].width = 116
        ws["A1"] = "Methodology, reconciliation and known limitations"
        F(ws["A1"], size=14, bold=True, color=C_HDR)
        a = attribute(self.p, INDEX, self.d0, self.d1) or {}
        jt = self.p.nearest_extra(self.d1, INDEX, "YTD Change (%)")
        blocks = [
            ("Data", [
                "Your history workbook is opened read-only and never written to. Each run re-reads it,",
                f"merges any JPM downloads found in the folder, and writes the growing copy to {EXTENDED_CSV}",
                "in the identical column layout. Snapshot-only fields - durations, JPM's own attribution,",
                f"ratings - go to {EXTRAS_CSV} keyed by date, so your schema never changes shape.",
                "JPM downloads are identified by their first column being 'Bam Id', never by filename.",
            ]),
            ("Year to date", [
                f"YTD is measured from the last observation of {self.year - 1} ({self.d0}), not from the",
                "first print of January. Using the first January print understates the year by whatever",
                "moved over the turn.",
            ]),
            ("Return attribution", [
                "total = -D x d(spread) - D x d(Treasury) + (carry, roll and residual)",
                "",
                "The identity closes by construction: the residual is defined as the remainder, not",
                "fitted. It is realised coupon, roll-down, rebalancing and defaults.",
                "",
                f"On the current YTD the residual is {a.get('residual', 0):+.2f}% against an index YTM of",
                f"{self.yld:.2f}%. That gap is not an error - it is what distressed names do to a quoted",
                "yield. Venezuela shows 42.8%, Lebanon 80.3%, Ethiopia -1.0%; none of them pay it. Any",
                "carry estimate built off the index YTM will be far too high.",
                "",
                "The Treasury leg uses the implied Treasury - yield less spread - which is the rate each",
                "entity is actually struck against, rather than an external 10y series.",
            ]),
            ("Duration", [
                "Attribution uses JPM's PUBLISHED spread duration from the nearest snapshot, because it",
                "cannot be estimated from the history: regressing daily total returns on yield changes",
                "returns 3.3 for the index against a published 6.22, for the same distressed-yield reason.",
                f"Where no snapshot covers a date the default of {DEFAULT_DURATION} is used and the column",
                "says 'default'. Duration is a sensitivity, not a maturity: the index shows 6.22 duration",
                "against 10.2y average life, and using average life would overstate every impact by 63%.",
            ]),
            ("Reconciliation", [
                (f"Our YTD total return {a.get('total', 0):+.3f}% against JPM's published {jt:+.3f}% - "
                 f"difference {(a.get('total', 0) - jt):+.3f}pp." if jt is not None else
                 "No JPM snapshot loaded, so no external reconciliation is available."),
                "Our spread and rate legs will NOT equal JPM's published legs: theirs are computed on the",
                "actual Treasury hedge and compound through the year, ours are a duration approximation at",
                "a single duration. Cite JPM's numbers externally; ours exist to run across every date in",
                "the history, which theirs cannot.",
            ]),
            ("Where the decomposition stops being informative", [
                "A duration split is a first-order expansion. It is accurate for small moves and",
                "fails for large ones, because duration itself changes as a credit moves. Rows are",
                f"flagged when the spread moved more than {LINEAR_LIMIT_BP:.0f}bp, when the residual",
                "exceeds the total, or when the index weight changed materially over the period.",
                "",
                "Credit C is the clearest case: -1,059bp of tightening, which at a duration of 4.95",
                "implies +52% against an actual +17.9%. Those names started the year with durations",
                "near 1-2 and the expansion has no chance. Read the total return there and ignore",
                "the legs. The same applies to the Middle East region, where the index weight fell",
                "from 16.6% to 12.2% and composition, not price, is doing much of the work.",
            ]),
            ("Term premium between the 10y and the index", [
                "The index is not priced off the 10y. Its spread is struck at each bond's own",
                "maturity and the cap-weighted average life is 10.2 years, so the Treasury it is",
                f"exposed to sits further out: {UST10Y_NOW}% on the 10y against an index anchor of",
                "yield less spread. The gap is term premium for the extra maturity and is close to",
                "what the cross-sectional curve implies at 4.9bp per year of average life.",
                "",
                "Duration and average life answer different questions and both are needed. Duration",
                "(6.22) sizes the P&L: 10bp is worth 0.62%. Average life (10.2y) picks WHICH rate,",
                "which is why the 10y is the reference and the 5y is not. A single 10.2y bond with",
                "this coupon would carry 7.4 duration; the index is shorter because duration is",
                "concave in maturity and the index barbells short distressed paper (Iraq 0.9y,",
                "Ethiopia 1.3y) against long investment grade (Peru 17.2y, Costa Rica 15.1y).",
                "",
                f"CURVE_BETA is {CURVE_BETA:.2f}: the anchor moves one-for-one with the 10y, so the",
                "premium cancels out of every change and the adjustment is presentational. It only",
                "bites if you model a flattening or steepening, and it cannot be estimated from",
                "these files because none of them carries a 10y series.",
            ]),
            ("Forecast", [
                "Betas are practitioner priors, not regression estimates.",
                "The known weakness: the rates beta is unconditional and cannot distinguish a rate rise",
                "driven by growth (spreads tighten) from one driven by an inflation or policy shock",
                "(spreads widen). If the call depends on which, build it regime by regime.",
                "Percentile bands are Gaussian around the central case and ignore fat tails.",
                "Two sigmas are shown: daily scaled by root-252, and the distribution of actual 252-day",
                "changes. They diverge when spreads mean-revert; the second is the honest one for a 12m",
                "horizon and is the one driving the bands.",
                "Sensitivity grid B holds spread and rates independent, which is conservative - in a real",
                "risk-off the Treasury rally offsets part of the widening.",
            ]),
        ]
        r = 3
        for title, lines in blocks:
            L(ws.cell(row=r, column=1), title, fill=C_BAND)
            r += 1
            for ln in lines:
                ws.cell(row=r, column=1, value=ln)
                F(ws.cell(row=r, column=1), color=C_NOTE)
                r += 1
            r += 1

    def build(self) -> Workbook:
        self.cover()
        self.ytd()
        self._breakdown("By_Region", "By region", REGIONS, True)
        self._breakdown("By_Rating", "By rating bucket", RATINGS, True)
        self._breakdown("By_Country", "By country", self.countries, True)
        self.spreads_tab()
        self.yields_tab()
        self.total_return_tab()
        self.breaks_tab()
        self.forecast()
        self.sensitivity()
        self.scenario_tab()
        self.attribution_tab()
        self.cushion_tab()
        self.spread_history_tab()
        self.return_history_tab()
        self.methodology()
        return self.wb


# ===========================================================================
# MONTHLY POWERPOINT
# ===========================================================================
#
# Built from the same panel as the workbook - never from re-keyed numbers.
# Needs python-pptx (pip install python-pptx); if it is missing the deck is
# skipped with a message rather than the whole run failing.

# Blue, red, grey, black only. Arial throughout.
PPT_INK = (0x00, 0x00, 0x00)             # body text
PPT_BLUE = (0x00, 0x2B, 0x5C)            # headings, table headers, primary series
PPT_BLUE2 = (0x00, 0x33, 0xA0)           # secondary blue
PPT_RED = (0xC8, 0x10, 0x2E)             # negatives, warnings
PPT_GREY = (0x8C, 0x8C, 0x8C)            # secondary series
PPT_MUTE = (0x59, 0x59, 0x59)            # captions
PPT_PAPER, PPT_SOFT = (0xFF, 0xFF, 0xFF), (0xEF, 0xF2, 0xF6)
PPT_OK, PPT_MARG, PPT_NOTOK = (0xC9, 0xD9, 0xEE), (0xE3, 0xE3, 0xE3), (0xF2, 0xC4, 0xC9)
PPT_NOTOK_TXT = (0x9C, 0x0A, 0x1E)
PPT_BLUE = PPT_BLUE                       # retained name, no gold in the palette
PPT_DEEP = PPT_BLUE
PPT_GOOD = PPT_BLUE


def month_bounds(panel: Panel) -> Tuple[Optional[str], str, str]:
    """(one month back, latest, label). Trailing from the run date, not the
    calendar month end - a deck run on the 20th should cover the last month,
    not the three weeks since the 1st."""
    d1 = panel.dates()[-1]
    target = dt(d1) - timedelta(days=30)
    prior = [d for d in panel.dates() if dt(d) <= target]
    return (prior[-1] if prior else None), d1, dt(d1).strftime("%B %Y")


def monthly_highlights(panel: Panel, d0: Optional[str], d1: str,
                       breaks: Dict[str, Any], countries: List[str]) -> List[str]:
    """Bullets written from the data, not from a template with numbers dropped in."""
    out: List[str] = []
    if not d0:
        return ["Not enough history for a month-on-month comparison."]
    idx = attribute(panel, INDEX, d0, d1) or {}
    tot, sb = idx.get("total"), idx.get("spread_bp")
    sp = panel.get(d1, INDEX, M_STW)
    if tot is not None:
        word = "returned" if tot >= 0 else "lost"
        out.append(f"Over the past month EMBIGD {word} {abs(tot):.2f}%, with the index spread "
                   f"{'tighter' if (sb or 0) < 0 else 'wider'} by {abs(sb or 0):.0f}bp to {sp:.0f}bp.")
    sb2 = idx.get("spread_bp")
    ub2 = idx.get("ust_bp")
    if sb2 is not None and ub2 is not None:
        lead = "spreads" if abs(sb2) >= abs(ub2) else "Treasury yields"
        out.append(f"{lead.capitalize()} did most of the work: spreads {sb2:+.0f}bp against "
                   f"{ub2:+.0f}bp on the underlying Treasury.")
    # regions
    reg = []
    for e in REGIONS:
        a = attribute(panel, e, d0, d1)
        if a and a.get("total") is not None:
            reg.append((a["total"], e, a.get("spread_bp")))
    if reg:
        reg.sort(reverse=True)
        best, worst = reg[0], reg[-1]
        out.append(f"{best[1].replace(' Region','')} led at {best[0]:+.2f}%; "
                   f"{worst[1].replace(' Region','')} lagged at {worst[0]:+.2f}%.")
    # IG vs non-IG
    ig = attribute(panel, "Credit IG only", d0, d1)
    hy = attribute(panel, "Credit Non-IG", d0, d1)
    if ig and hy and ig.get("total") is not None and hy.get("total") is not None:
        gap = hy["total"] - ig["total"]
        out.append(f"Non-IG {'outperformed' if gap > 0 else 'underperformed'} investment grade by "
                   f"{abs(gap):.2f}pp ({hy['total']:+.2f}% against {ig['total']:+.2f}%).")
    # movers
    cr = []
    for e in countries:
        a = attribute(panel, e, d0, d1)
        w = panel.get(d1, e, M_WGT) or 0
        if a and a.get("total") is not None and w > 0.15:
            cr.append((a["total"], e))
    if len(cr) >= 6:
        cr.sort(reverse=True)
        up = ", ".join(f"{e} {t:+.1f}%" for t, e in cr[:3])
        dn = ", ".join(f"{e} {t:+.1f}%" for t, e in cr[-3:][::-1])
        out.append(f"Best: {up}.")
        out.append(f"Worst: {dn}.")
    # breaks inside the month
    inside = [d for d in breaks if d0 < d <= d1]
    if inside:
        d = inside[-1]
        out.append(f"NOTE: JPM changed the spread basis on {d} — the series stepped "
                   f"{breaks[d]['jump_bp']:+.0f}bp with no matching return. Spread moves here are "
                   f"chain-linked across it and measure the market, not the methodology.")
    return out


def build_deck(panel: Panel, dash: "Dashboard", path: Path) -> bool:
    try:
        from pptx import Presentation
        from pptx.chart.data import CategoryChartData
        from pptx.dml.color import RGBColor
        from pptx.enum.chart import XL_CHART_TYPE, XL_LEGEND_POSITION
        from pptx.enum.text import PP_ALIGN
        from pptx.util import Emu, Inches, Pt
    except ImportError:
        print("Deck     : skipped — python-pptx not installed (pip install python-pptx)",
              file=sys.stderr)
        return False

    d0, d1, label = month_bounds(panel)
    prs = Presentation()
    prs.slide_width, prs.slide_height = Inches(13.333), Inches(7.5)
    blank = prs.slide_layouts[6]

    def rgb(t):
        return RGBColor(*t)

    def bp(v) -> str:
        return "n/a" if v is None else f"{v:+.0f}"

    def txt(sl, x, y, w, h, text, size=14, bold=False, color=PPT_INK,
            font="Arial", align=PP_ALIGN.LEFT, italic=False):
        tb = sl.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
        tf = tb.text_frame
        tf.word_wrap = True
        tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
        p = tf.paragraphs[0]
        p.alignment = align
        r = p.add_run()
        r.text = text
        r.font.size, r.font.bold, r.font.italic = Pt(size), bold, italic
        r.font.color.rgb = rgb(color)
        r.font.name = font
        return tb

    def bullets(sl, x, y, w, h, items, size=13):
        tb = sl.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
        tf = tb.text_frame
        tf.word_wrap = True
        for i, it in enumerate(items):
            p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
            p.space_after = Pt(10)
            r = p.add_run()
            r.text = "•  " + it
            r.font.size = Pt(size)
            r.font.name = "Arial"
            r.font.color.rgb = rgb(PPT_RED if it.startswith("NOTE:") else PPT_INK)
            r.font.bold = it.startswith("NOTE:")
        return tb

    def style_chart(gf, colours, numfmt='+0.0;-0.0', label_size=9):
        """Category labels pinned low so a negative bar never sits on top of
        them, and series in blue/grey/red only."""
        gf.has_title = False
        ca = gf.category_axis
        # python-pptx writes tickLblPos="nextTo" on a new chart and the setter
        # does not always stick, so set the element itself. "low" pins the
        # category labels to the bottom of the plot area, which is what keeps
        # them readable when a bar runs negative.
        try:
            el = ca._element.find(
                '{http://schemas.openxmlformats.org/drawingml/2006/chart}tickLblPos')
            if el is not None:
                el.set("val", "low")
        except Exception:
            pass
        ca.tick_labels.font.size = Pt(11)
        ca.tick_labels.font.name = "Arial"
        ca.tick_labels.font.color.rgb = rgb(PPT_INK)
        va = gf.value_axis
        va.tick_labels.font.size = Pt(10)
        va.tick_labels.font.name = "Arial"
        va.tick_labels.font.color.rgb = rgb(PPT_MUTE)
        va.has_major_gridlines = True
        pl = gf.plots[0]
        pl.has_data_labels = True
        pl.data_labels.number_format = numfmt
        pl.data_labels.number_format_is_linked = False
        pl.data_labels.font.size = Pt(label_size)
        pl.data_labels.font.name = "Arial"
        pl.data_labels.font.color.rgb = rgb(PPT_INK)
        for k, col in enumerate(colours):
            if k < len(pl.series):
                f_ = pl.series[k].format.fill
                f_.solid(); f_.fore_color.rgb = rgb(col)
        return gf

    def fill(sl, x, y, w, h, color):
        from pptx.enum.shapes import MSO_SHAPE
        sh = sl.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(x), Inches(y),
                                 Inches(w), Inches(h))
        sh.fill.solid(); sh.fill.fore_color.rgb = rgb(color)
        sh.line.fill.background()
        sh.shadow.inherit = False
        sh.text_frame.text = ""
        return sh

    def table(sl, x, y, w, rows, widths, size=12, header=True):
        nr, nc = len(rows), len(rows[0])
        h = 0.32 * nr
        gt = sl.shapes.add_table(nr, nc, Inches(x), Inches(y), Inches(w), Inches(h)).table
        for j, cw in enumerate(widths):
            gt.columns[j].width = Emu(int(Inches(cw)))
        for i, row in enumerate(rows):
            for j, val in enumerate(row):
                c = gt.cell(i, j)
                c.text = str(val)
                p = c.text_frame.paragraphs[0]
                p.alignment = PP_ALIGN.LEFT if j == 0 else PP_ALIGN.RIGHT
                for r_ in p.runs:
                    r_.font.size = Pt(size)
                    r_.font.name = "Arial"
                    r_.font.bold = (i == 0 and header)
                    r_.font.color.rgb = rgb(PPT_PAPER if (i == 0 and header) else PPT_INK)
                c.fill.solid()
                c.fill.fore_color.rgb = rgb(PPT_BLUE if (i == 0 and header)
                                            else (PPT_SOFT if i % 2 else PPT_PAPER))
        return gt

    idx = attribute(panel, INDEX, d0, d1) if d0 else {}
    ytd = attribute(panel, INDEX, dash.d0, d1) or {}
    sp = panel.get(d1, INDEX, M_STW) or 0

    # ---------- 1. headline ----------
    s1 = prs.slides.add_slide(blank)
    txt(s1, 0.7, 0.6, 12, 0.9, f"EMBI Global Diversified — {label}", 34, True,
        PPT_BLUE, "Arial")
    txt(s1, 0.7, 1.45, 12, 0.4,
        f"Month to {d1}" + (f", against {d0}" if d0 else ""), 14, False, PPT_MUTE)
    tiles = [(f"{(idx or {}).get('total', 0):+.2f}%", "total return on the month", PPT_PAPER),
             (f"{sp:.0f}bp", "index spread", PPT_PAPER),
             (f"{(idx or {}).get('spread_bp', 0):+.0f}bp", "spread move on the month", PPT_BLUE),
             (f"{ytd.get('total', 0):+.2f}%", f"total return YTD {dash.year}", PPT_BLUE)]
    for i, (v, l, c) in enumerate(tiles):
        x = 0.7 + i * 3.05
        fill(s1, x, 2.2, 2.8, 1.65, PPT_SOFT)
        neg = v.startswith("-")
        txt(s1, x, 2.4, 2.8, 0.7, v, 32, True, PPT_RED if neg else PPT_BLUE,
            "Arial", PP_ALIGN.CENTER)
        txt(s1, x + 0.15, 3.15, 2.5, 0.6, l, 11, False, PPT_MUTE, "Arial", PP_ALIGN.CENTER)
    # One labelled column per period. A single sentence carrying four numbers
    # invites the reader to attach the last one to the first label.
    periods = [("Past month", d0), ("3 months", dash._back(3)),
               (f"YTD {dash.year}", dash.d0), ("12 months", dash._back(12))]
    hdr_row, ret_row, spr_row = ["Period"], ["Total return %"], ["Spread change bp"]
    for lbl, dfrom in periods:
        tr_ = dash._chg(INDEX, M_TRI, dfrom, d1)
        sb_ = dash._chg(INDEX, M_STW, dfrom, d1)
        hdr_row.append(lbl)
        ret_row.append(f"{tr_:+.2f}" if tr_ is not None else "—")
        spr_row.append(f"{sb_:+.0f}" if sb_ is not None else "—")
    txt(s1, 0.7, 4.2, 11.9, 0.3, "EMBIGD total return, from the index's own return series",
        13, True, PPT_BLUE)
    table(s1, 0.7, 4.55, 9.6, [hdr_row, ret_row, spr_row],
          [2.4, 1.8, 1.8, 1.8, 1.8], size=12)
    hist = [(d, panel.get(d, INDEX, M_STW)) for d in dash.month_ends(13)]
    hist = [(d, v_) for d, v_ in hist if v_ is not None]
    if len(hist) >= 4:
        txt(s1, 0.7, 5.75, 11.9, 0.3, "Index spread, last twelve months (bp)", 13, True, PPT_BLUE)
        cd2 = CategoryChartData()
        cd2.categories = [d[:7] for d, _ in hist]
        cd2.add_series("Spread", [v_ for _, v_ in hist])
        gf2 = s1.shapes.add_chart(XL_CHART_TYPE.LINE, Inches(0.7), Inches(6.05),
                                  Inches(11.9), Inches(0.85), cd2).chart
        gf2.has_legend = False
        gf2.has_title = False
        ca2 = gf2.category_axis
        ca2.tick_labels.font.size = Pt(8)
        ca2.tick_labels.font.name = "Arial"
        ca2.tick_labels.font.color.rgb = rgb(PPT_MUTE)
        va2 = gf2.value_axis
        va2.tick_labels.font.size = Pt(8)
        va2.tick_labels.font.name = "Arial"
        va2.tick_labels.font.color.rgb = rgb(PPT_MUTE)
        ln = gf2.plots[0].series[0].format.line
        ln.color.rgb = rgb(PPT_BLUE); ln.width = Pt(2)
    txt(s1, 0.7, 7.0, 11.9, 0.3,
        "Source: JPM EMBI Global Diversified. Spread changes chain-linked across basis breaks; "
        "the level series is shown as published.", 9, False, PPT_MUTE)

    # ---------- 2. by region: past month AND year to date ----------
    s2 = prs.slides.add_slide(blank)
    txt(s2, 0.7, 0.5, 12, 0.7, "Performance by region", 34, True, PPT_INK, "Arial")
    txt(s2, 0.7, 1.2, 12, 0.35,
        f"Past month {d0} to {d1}  ·  year to date from {dash.d0}", 14, False, PPT_MUTE)
    reg = []
    for e in REGIONS:
        m = attribute(panel, e, d0, d1) if d0 else None
        y = attribute(panel, e, dash.d0, d1)
        if m and y and m.get("total") is not None and y.get("total") is not None:
            reg.append((e.replace(" Region", ""), m, y, e))
    reg.sort(key=lambda t: -t[2]["total"])
    if reg:
        cd = CategoryChartData()
        cd.categories = [e for e, _m, _y, _f in reg]
        cd.add_series("Past month %", [m["total"] for _e, m, _y, _f in reg])
        cd.add_series("Year to date %", [y["total"] for _e, _m, y, _f in reg])
        gf = s2.shapes.add_chart(XL_CHART_TYPE.COLUMN_CLUSTERED, Inches(0.7), Inches(1.75),
                                 Inches(6.3), Inches(4.3), cd).chart
        gf.has_legend = True
        gf.legend.position = XL_LEGEND_POSITION.BOTTOM
        gf.legend.include_in_layout = False
        gf.legend.font.size = Pt(11)
        gf.legend.font.name = "Arial"
        style_chart(gf, [PPT_GREY, PPT_BLUE])
        rows = [["Region", "Wt %", "Spread", "1m bp", "1m %", "YTD bp", "YTD %"]]
        for e, m, y, full in reg:
            rows.append([e, f"{panel.get(d1, full, M_WGT) or 0:.1f}",
                         f"{panel.get(d1, full, M_STW) or 0:.0f}",
                         bp(m.get('spread_bp')), f"{m['total']:+.2f}",
                         bp(y.get('spread_bp')), f"{y['total']:+.2f}"])
        idxm = attribute(panel, INDEX, d0, d1) if d0 else None
        idxy = attribute(panel, INDEX, dash.d0, d1)
        if idxm and idxy:
            rows.append(["EMBIGD", "100.0", f"{sp:.0f}",
                         f"{idxm.get('spread_bp', 0):+.0f}", f"{idxm['total']:+.2f}",
                         f"{idxy.get('spread_bp', 0):+.0f}", f"{idxy['total']:+.2f}"])
        table(s2, 7.3, 1.9, 5.3, rows, [1.25, 0.6, 0.75, 0.65, 0.7, 0.7, 0.7], size=11)
    txt(s2, 7.3, 5.05, 5.3, 1.5,
        "Spread moves are chain-linked, so JPM's 04-Sep exclusion of defaulted names does not "
        "appear here as a market move. Returns need no such treatment. Note that spread level "
        "minus spread change will not tie across that date — the level is on the new basis and "
        "the change is the market's.",
        11, False, PPT_MUTE, italic=True)
    txt(s2, 0.7, 6.85, 11.9, 0.3, f"Source: JPM EMBI Global Diversified, {d1}.",
        9, False, PPT_MUTE)

    # ---------- 3. highlights ----------
    s3 = prs.slides.add_slide(blank)
    txt(s3, 0.7, 0.5, 12, 0.7, f"{label} — highlights", 34, True, PPT_INK, "Arial")
    hl = monthly_highlights(panel, d0, d1, dash.breaks, dash.countries)
    bullets(s3, 0.7, 1.5, 7.4, 4.8, hl, 13)
    rows = [["Rating bucket", "Wt %", "Spread", "1m %", "YTD %"]]
    for e in ["Credit IG only", "Credit BB only", "Credit B only", "Credit C only",
              "Credit Non-IG"]:
        m = attribute(panel, e, d0, d1) if d0 else None
        y = attribute(panel, e, dash.d0, d1)
        w = panel.get(d1, e, M_WGT)
        if not m or not y or m.get("total") is None or not w:
            continue
        rows.append([e.replace("Credit ", "").replace(" only", ""), f"{w:.1f}",
                     f"{panel.get(d1, e, M_STW) or 0:.0f}",
                     f"{m['total']:+.2f}", f"{y['total']:+.2f}"])
    if len(rows) > 1:
        table(s3, 8.3, 1.6, 4.3, rows, [1.35, 0.65, 0.8, 0.75, 0.75], size=11)
    txt(s3, 0.7, 6.85, 11.9, 0.3, "Highlights generated from the underlying series, not re-keyed.",
        9, False, PPT_MUTE)

    # ---------- 3b. who is killing it ----------
    s5 = prs.slides.add_slide(blank)
    txt(s5, 0.7, 0.5, 12, 0.7, "Where the performance came from", 34, True, PPT_INK, "Arial")
    txt(s5, 0.7, 1.2, 12, 0.35,
        "Countries above 0.25% of the index, ranked by total return. Past month on the left, "
        "year to date on the right.", 14, False, PPT_MUTE)

    def ranked(dfrom: Optional[str]) -> List[Tuple[float, str, Any]]:
        out_ = []
        for e in dash.countries:
            w = panel.get(d1, e, M_WGT) or 0
            if w < 0.25:
                continue
            at = attribute(panel, e, dfrom, d1) if dfrom else None
            if at and at.get("total") is not None:
                out_.append((at["total"], e, at))
        out_.sort(reverse=True)
        return out_

    for col, (dfrom, head) in enumerate([(d0, "Past month"), (dash.d0, "Year to date")]):
        x = 0.7 + col * 6.3
        rk = ranked(dfrom)
        if not rk:
            continue
        txt(s5, x, 1.8, 5.4, 0.3, head, 14, True, PPT_BLUE)
        rows = [["Leaders", "Wt %", "Spread bp", "Return %"]]
        for t, e, at in rk[:7]:
            rows.append([e[:20], f"{panel.get(d1, e, M_WGT) or 0:.1f}",
                         bp(at.get('spread_bp')), f"{t:+.2f}"])
        table(s5, x, 2.3, 5.8, rows, [2.6, 0.9, 1.15, 1.15], size=11)
        rows = [["Laggards", "Wt %", "Spread bp", "Return %"]]
        for t, e, at in rk[-4:][::-1]:
            rows.append([e[:20], f"{panel.get(d1, e, M_WGT) or 0:.1f}",
                         bp(at.get('spread_bp')), f"{t:+.2f}"])
        table(s5, x, 4.9, 5.8, rows, [2.6, 0.9, 1.15, 1.15], size=11)
    txt(s5, 0.7, 6.85, 11.9, 0.3,
        "Spread changes chain-linked; names below 0.25% of the index are excluded so the list "
        "is not dominated by positions too small to trade.", 9, False, PPT_MUTE)

    # ---------- 4. projections ----------
    s4 = prs.slides.add_slide(blank)
    txt(s4, 0.7, 0.5, 12, 0.7, "Projections — US 10-year scenarios", 34, True, PPT_INK, "Arial")
    txt(s4, 0.7, 1.2, 12, 0.4,
        f"US 10y at {dash.ust10y:.2f}%. The index is struck at 10.2y average life, so the "
        f"Treasury it is exposed to sits {dash.term_premium * 100:+.0f}bp further out at "
        f"{dash.ust_now:.2f}%. Spread beta {BETAS['ust10y_bp']:.2f}, duration {dash.D:.2f}, "
        f"cash {CASH_RATE:.2f}%.", 13, False, PPT_MUTE)
    rows = [["US 10y", "Index anchor", "Spread (bp)", "12m return %", "vs cash (pp)"]]
    cats, vals = [], []
    for lvl in dash.ust_grid:
        u = (lvl - dash.ust10y) * 100.0
        anchor = lvl + dash.term_premium + (CURVE_BETA - 1.0) * u / 100.0
        du = (anchor - dash.ust_now) * 100.0
        s_ = sp + BETAS["ust10y_bp"] * u
        tr = total_return(dash.yld, dash.IRD, dash.D, du, sp, s_)
        rows.append([f"{lvl:.2f}%", f"{anchor:.2f}%", f"{s_:.0f}", f"{tr:.2f}",
                     f"{tr - CASH_RATE:+.2f}"])
        cats.append(f"{lvl:.2f}%")
        vals.append(tr)
    table(s4, 0.7, 1.95, 5.6, rows, [1.05, 1.2, 1.15, 1.15, 1.05], size=11)
    cd = CategoryChartData()
    cd.categories = cats
    cd.add_series("12m total return %", vals)
    gf = s4.shapes.add_chart(XL_CHART_TYPE.COLUMN_CLUSTERED, Inches(6.2), Inches(1.9),
                             Inches(6.4), Inches(3.6), cd).chart
    gf.has_legend = False
    style_chart(gf, [PPT_BLUE], numfmt='0.00', label_size=10)
    be = sp + (dash.yld - CASH_RATE) * 100.0 / dash.D
    txt(s4, 0.7, 5.85, 11.9, 0.9,
        f"Break-even: the call stops beating cash at {be:.0f}bp with the 10y at "
        f"{dash.ust10y:.2f}%. "
        f"Scenarios hold spread and rates linked through beta; in a genuine risk-off the "
        f"Treasury rally offsets part of the widening, so this is the conservative read.",
        12, False, PPT_INK, italic=True)
    txt(s4, 0.7, 6.85, 11.9, 0.3,
        f"Curve beta {CURVE_BETA:.2f} (parallel shift). Betas are practitioner priors, not "
        f"fitted; the rates beta is unverified — see the workbook's Methodology tab.",
        9, False, PPT_MUTE)

    # ---------- 5. scenario grid ----------
    s5 = prs.slides.add_slide(blank)
    txt(s5, 0.7, 0.5, 12, 0.7, "Scenario grid — total return by US 10-year and spread level",
        30, True, PPT_INK, "Arial")
    txt(s5, 0.7, 1.15, 12, 0.45,
        f"12-month total return, %. Rows are index spread levels, columns the US 10y. Spot: "
        f"spread {sp:.0f}bp, 10y {dash.ust10y:.2f}%, yield {dash.yld:.2f}%, duration {dash.D:.2f}. "
        f"Blue beats cash ({CASH_RATE:.2f}%) by {GRID_MARGIN_PP:.0f}pp or more, grey is within "
        f"{GRID_MARGIN_PP:.0f}pp of cash, red is below cash by more than {GRID_MARGIN_PP:.0f}pp.",
        12, False, PPT_MUTE)
    levels, grows = scenario_grid(dash)
    nr, nc = len(grows) + 1, len(levels) + 1
    gw = 12.6
    gt = s5.shapes.add_table(nr, nc, Inches(0.7), Inches(1.75), Inches(gw),
                             Inches(0.36 * nr)).table
    first_w = 1.35
    gt.columns[0].width = Emu(int(Inches(first_w)))
    for j in range(1, nc):
        gt.columns[j].width = Emu(int(Inches((gw - first_w) / len(levels))))
    spot_j = None
    for j, lvl in enumerate(levels, start=1):
        if abs(lvl - dash.ust10y) <= 0.005:
            spot_j = j

    def gcell(i, j, text, fill_c, txt_c, bold=False, align=PP_ALIGN.RIGHT):
        c = gt.cell(i, j)
        c.text = text
        c.margin_left = c.margin_right = Emu(int(Inches(0.05)))
        p = c.text_frame.paragraphs[0]
        p.alignment = align
        for r_ in p.runs:
            r_.font.size = Pt(10); r_.font.name = "Arial"
            r_.font.bold = bold; r_.font.color.rgb = rgb(txt_c)
        c.fill.solid(); c.fill.fore_color.rgb = rgb(fill_c)

    gcell(0, 0, "Spread \\ US 10y", PPT_BLUE, PPT_PAPER, True, PP_ALIGN.LEFT)
    for j, lvl in enumerate(levels, start=1):
        gcell(0, j, f"{lvl:.2f}%" + (" spot" if j == spot_j else ""), PPT_BLUE, PPT_PAPER, True)
    for i, (lab, s_, vals) in enumerate(grows, start=1):
        spot_row = lab.startswith("Spot")
        gcell(i, 0, lab, PPT_SOFT if spot_row else PPT_PAPER, PPT_INK, spot_row, PP_ALIGN.LEFT)
        for j, tr in enumerate(vals, start=1):
            v = grid_verdict(tr)
            fill_c = {"ok": PPT_OK, "marginal": PPT_MARG, "bad": PPT_NOTOK}[v]
            txt_c = {"ok": PPT_BLUE, "marginal": PPT_INK, "bad": PPT_NOTOK_TXT}[v]
            gcell(i, j, f"{tr:.1f}", fill_c, txt_c, bold=(j == spot_j or spot_row))
    # break-even line and where the competing forecasts land
    be_now = sp + (dash.yld - CASH_RATE) * 100.0 / dash.D
    be_up = sp + (dash.yld - dash.IRD * 0.5 - CASH_RATE) * 100.0 / dash.D
    be_dn = sp + (dash.yld + dash.IRD * 0.5 - CASH_RATE) * 100.0 / dash.D
    y_note = 1.75 + 0.36 * nr + 0.2
    txt(s5, 0.7, y_note, 11.9, 0.9,
        f"Break-even spread: {be_now:.0f}bp with the 10y unchanged at {dash.ust10y:.2f}%; "
        f"{be_up:.0f}bp if the 10y rises 50bp to {dash.ust10y + 0.5:.2f}%; {be_dn:.0f}bp if it "
        f"falls 50bp to {dash.ust10y - 0.5:.2f}%. Every basis point of spread is worth "
        f"{dash.D / 100:.3f}% of return; every basis point on the 10y is worth "
        f"{dash.IRD / 100:.3f}%.",
        12, False, PPT_INK, italic=True)
    txt(s5, 0.7, y_note + 0.95, 11.9, 0.3,
        "Spread and rates set independently — the bottom-right overstates the pain of a risk-off, "
        "where the Treasury rally offsets part of the widening. Convexity ignored.",
        9, False, PPT_MUTE)

    prs.save(str(path))
    return True


# ===========================================================================
# DRIVER
# ===========================================================================

def find_history_xlsx(folder: Path) -> Optional[Path]:
    best = None
    for p in folder.glob("*.xlsx"):
        if p.name.startswith("~$") or p.name == OUTPUT_XLSX:
            continue
        if any(h in p.name.lower() for h in HISTORY_XLSX_HINTS):
            return p
        best = best or p
    return best


def main(argv: Optional[List[str]] = None) -> int:
    ap = argparse.ArgumentParser(description="EMBI Global Diversified tracker.")
    ap.add_argument("--dir", default=".", help="folder to work in (default: current)")
    ap.add_argument("--history", default=None, help="path to the history workbook")
    ap.add_argument("--output", "-o", default=OUTPUT_XLSX)
    ap.add_argument("--no-build", action="store_true", help="ingest only")
    ap.add_argument("--check", action="store_true", help="reconciliation report only")
    ap.add_argument("--no-deck", action="store_true", help="skip the monthly PowerPoint")
    ap.add_argument("--deck", default=None, help="PowerPoint filename (default: EMBI_Monthly_<month>.pptx)")
    a = ap.parse_args(argv)

    folder = Path(a.dir).expanduser().resolve()
    panel = Panel()

    hist = Path(a.history).expanduser() if a.history else find_history_xlsx(folder)
    if hist and hist.exists():
        n = panel.load_history_xlsx(hist)
        print(f"History  : {hist.name} — {n:,} rows (read-only, untouched)")
    else:
        print("History  : none found — starting from snapshots alone", file=sys.stderr)

    n_ext = panel.load_extended_csv(folder / EXTENDED_CSV)
    if n_ext:
        print(f"Extended : {EXTENDED_CSV} — {n_ext:,} rows")
    panel.load_extras_csv(folder / EXTRAS_CSV)

    known = panel.entities()
    archive = folder / ARCHIVE_DIR
    files = sorted(folder.glob("*.csv"))
    if archive.exists():
        files += sorted(archive.glob("*.csv"))
    added = 0
    rejected = 0
    for p in files:
        if p.name in (EXTENDED_CSV, EXTRAS_CSV) or not Panel.is_snapshot(p):
            continue
        res = panel.ingest_snapshot(p, known)
        if not res:
            rejected += 1
            continue
        d, nv, new = res
        print(f"Snapshot : {d}  {p.name}  ({nv} values, {new} new series)")
        added += 1
        archive.mkdir(exist_ok=True)
        dest = archive / f"snapshot_{d}.csv"
        if p.resolve() != dest.resolve() and not dest.exists():
            dest.write_bytes(p.read_bytes())

    if rejected:
        print(f"\n{rejected} file(s) REJECTED as not EMBI Global Diversified — see above. If any "
              f"of them sit in {ARCHIVE_DIR}/, delete them there too, and delete "
              f"{EXTENDED_CSV} and {EXTRAS_CSV} so a clean copy is rebuilt from the history "
              f"workbook on this run.", file=sys.stderr)
    if not panel.data:
        print("ERROR: no data at all.", file=sys.stderr)
        return 1

    dates = panel.dates()
    print(f"\nPanel    : {len(dates):,} dates, {dates[0]} to {dates[-1]}, "
          f"{len(panel.entities())} entities")

    panel._breaks = detect_breaks(panel)
    if panel._breaks:
        print(f"\nBasis breaks detected on {len(panel._breaks)} day(s) — spread stepped, "
              f"return did not:")
        for d in sorted(panel._breaks):
            b = panel._breaks[d]
            print(f"   {d}  spread {b['jump_bp']:+.0f}bp  but return {b['actual_pct']:+.2f}% "
                  f"(the jump implied {b['implied_pct']:+.2f}%)")
        print("   All period changes below are CHAIN-LINKED across these days.")

    d1 = dates[-1]
    d0 = year_start(panel, dt(d1).year) or dates[0]
    at = attribute(panel, INDEX, d0, d1) or {}
    D, dsrc = panel.duration(d1, INDEX)
    print(f"\nYTD {dt(d1).year} ({d0} -> {d1}), duration {D:.2f} ({dsrc})")
    print(f"   total return              {at.get('total', 0):+7.2f}%")
    print(f"   spread move (chain-linked){at.get('spread_bp', 0):+7.0f}bp")
    if at.get("crossed_break"):
        print(f"   spread move (naive)       {at.get('spread_bp_naive', 0):+7.0f}bp"
              f"   <- crosses a basis break, do not use")
    print(f"   underlying Treasury move  {at.get('ust_bp', 0):+7.0f}bp")
    print(f"   worth of the spread move  {at.get('spread_leg', 0):+7.2f}%")
    print(f"   worth of the rate move    {at.get('rate_leg', 0):+7.2f}%")
    print(f"   carry, roll and residual  {at.get('residual', 0):+7.2f}%")
    jt = panel.nearest_extra(d1, INDEX, "YTD Change (%)")
    if jt is not None:
        diff = (at.get("total") or 0) - jt
        flag = "OK" if abs(diff) < 0.02 else "CHECK"
        print(f"   vs JPM published total    {jt:+7.2f}%   diff {diff:+.3f}pp  [{flag}]")

    if a.check:
        return 0

    panel.save_extended(folder / EXTENDED_CSV)
    panel.save_extras(folder / EXTRAS_CSV)
    print(f"\nWrote    : {EXTENDED_CSV} ({len(dates):,} rows)  and  {EXTRAS_CSV}")
    if added:
        print(f"           {added} snapshot(s) archived to {ARCHIVE_DIR}/")

    if a.no_build:
        return 0
    out = Path(a.output)
    if not out.is_absolute():
        out = folder / out
    dash = Dashboard(panel)
    dash.build().save(out)
    print(f"Wrote    : {out.name}")

    if not a.no_deck:
        _d0, d1x, label = month_bounds(panel)
        name = a.deck or f"EMBI_Monthly_{dt(d1x).strftime('%Y_%m')}.pptx"
        dpath = Path(name)
        if not dpath.is_absolute():
            dpath = folder / dpath
        if build_deck(panel, dash, dpath):
            print(f"Wrote    : {dpath.name}  ({label} monthly deck)")
    return 0


if __name__ == "__main__":
    sys.exit(main())
