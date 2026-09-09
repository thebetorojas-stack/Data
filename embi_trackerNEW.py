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
SENSITIVITY_UST = [-100, -75, -50, -25, 0, 25, 50, 75, 100]
SENSITIVITY_SPREADS = [125, 150, 175, 200, 225, 250, 275]

# A duration approximation is a first-order expansion: it is accurate for small
# moves and fails badly for large ones, because duration itself changes as the
# credit moves. Beyond this many bp the attribution is flagged rather than
# quietly presented. Credit C moved -1059bp this year: -D x ds implies +52%
# against an actual +18%, so the legs there are arithmetic, not analysis.
LINEAR_LIMIT_BP = 200.0

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
SNAP_ENTITY_MAP = {
    "embi global diversified": INDEX, "embi global": INDEX, "embig div": INDEX,
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
C_HDR, C_HDRFG = "1F4E78", "FFFFFF"
C_BAND, C_GOOD, C_BAD = "D9E1F2", "E2EFDA", "FCE4E4"
C_IN, C_BRD = "FFF2CC", "BFBFBF"
C_HARD, C_TXT, C_NOTE = "0000FF", "000000", "7F7F7F"
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
                ent = INDEX          # the first data row is always the index
                first = False
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


# ===========================================================================
# ANALYTICS
# ===========================================================================

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
    out["spread_bp"] = s1 - s0
    out["dur"], out["dur_src"] = D, src
    out["spread_leg"] = -D * (s1 - s0) / 100.0
    if u0 is not None and u1 is not None:
        out["ust_bp"] = (u1 - u0) * 100.0
        out["rate_leg"] = -D * (u1 - u0)
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
    if abs(out["spread_bp"] or 0) > LINEAR_LIMIT_BP:
        reasons.append(f"spread moved {abs(out['spread_bp']):.0f}bp — too large for a duration split")
    if out["residual"] is not None and out["residual"] < -1.0:
        reasons.append("negative carry — defaults or composition, not coupon")
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


def total_return(yld: float, ir_dur: float, spr_dur: float,
                 d_ust_bp: float, spread_now: float, spread_fcst: float) -> float:
    return yld - ir_dur * (d_ust_bp / 100.0) - spr_dur * ((spread_fcst - spread_now) / 100.0)


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
        self.wb = Workbook()
        self.wb.remove(self.wb.active)
        self.D, self.dsrc = panel.duration(self.d1, INDEX)
        self.IRD = panel.ir_duration(self.d1, INDEX)
        self.yld = panel.get(self.d1, INDEX, M_YLD) or 0.0
        self.spread = panel.get(self.d1, INDEX, M_STW) or 0.0
        self.countries = [e for e in panel.entities()
                          if e != INDEX and e not in REGIONS and e not in RATINGS]

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
        ws["A2"] = ("Total return split into what the spread move was worth, what the Treasury "
                    "move was worth, and the remainder — realised coupon, roll, rebalancing and "
                    "defaults. The three legs add to the total by construction.")
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
            ("", None, "", ""),
            ("  worth of the spread move", a.get("spread_leg"), "+0.00;-0.00",
             f"= -{self.D:.2f} x spread change"),
            ("  worth of the rate move", a.get("rate_leg"), "+0.00;-0.00",
             f"= -{self.D:.2f} x Treasury change"),
            ("  carry, roll and residual", a.get("residual"), "+0.00;-0.00",
             "the remainder, not an error"),
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
        ws["A2"] = ("Spread and rate legs are at published duration; 'carry & residual' is the "
                    "remainder that makes the three add to the total. Shaded rows are flagged in "
                    "the last column — the decomposition there is arithmetic, not analysis.")
        F(ws["A2"], italic=True, color=C_NOTE)
        cols = [("", 30), ("Weight %", 10), ("Spread now", 11), ("Spread YTD bp", 13),
                ("Treasury YTD bp", 14), ("Yield %", 9), ("Duration", 10),
                ("YTD total %", 12), ("of which spread", 15), ("of which rates", 14),
                ("carry & residual", 16), ("spread share", 12), ("read with care", 34)]
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
        r = hr + 1
        for _k, e, at, w in rows:
            L(ws.cell(row=r, column=1), e, bold=False)
            V(ws.cell(row=r, column=2), w, fmt="0.00", color=C_HARD)
            V(ws.cell(row=r, column=3), self.p.get(self.d1, e, M_STW), fmt="0", color=C_HARD)
            V(ws.cell(row=r, column=4), at.get("spread_bp"), fmt="+0;-0")
            V(ws.cell(row=r, column=5), at.get("ust_bp"), fmt="+0;-0")
            V(ws.cell(row=r, column=6), self.p.get(self.d1, e, M_YLD), fmt="0.00", color=C_HARD)
            V(ws.cell(row=r, column=7), at.get("dur"), fmt="0.00")
            V(ws.cell(row=r, column=8), at.get("total"), fmt="+0.00;-0.00", bold=True, color=C_HARD)
            V(ws.cell(row=r, column=9), at.get("spread_leg"), fmt="+0.00;-0.00")
            V(ws.cell(row=r, column=10), at.get("rate_leg"), fmt="+0.00;-0.00")
            V(ws.cell(row=r, column=11), at.get("residual"), fmt="+0.00;-0.00")
            tot, sl = at.get("total"), at.get("spread_leg")
            share = (sl / tot) if (tot and abs(tot) >= 0.25 and sl is not None) else None
            if share is None:
                V(ws.cell(row=r, column=12), "n/m", fmt="General", color=C_NOTE)
            else:
                V(ws.cell(row=r, column=12), share, fmt="0%")
            flag = at.get("flag") or ""
            c = V(ws.cell(row=r, column=13), flag, fmt="General", color=C_NOTE)
            c.alignment = Alignment(horizontal="left")
            F(c, italic=True, color=C_NOTE)
            if flag:
                for cc in range(8, 13):
                    ws.cell(row=r, column=cc).fill = PatternFill("solid", fgColor=C_BAD)
            r += 1
        ws.freeze_panes = "B5"

    # ---------- history ----------
    def history_tab(self):
        ws = self.wb.create_sheet("Spread_History")
        ws.sheet_view.showGridLines = False
        ws["A1"] = "Spread history — month ends (STW Trsy, bp)"
        F(ws["A1"], size=14, bold=True, color=C_HDR)
        ws["A2"] = "Every observation is kept in the extended CSV; this tab samples month ends for legibility."
        F(ws["A2"], italic=True, color=C_NOTE)
        month_end: Dict[str, str] = {}
        for d in self.dates:
            month_end[d[:7]] = d
        picks = [month_end[k] for k in sorted(month_end)]
        if self.d1 not in picks:
            picks.append(self.d1)
        picks = picks[-60:]
        hr = 4
        H(ws.cell(row=hr, column=1), "Entity")
        ws.column_dimensions["A"].width = 30
        for j, d in enumerate(picks, start=2):
            H(ws.cell(row=hr, column=j), d[:7] if d != self.d1 else d)
            ws.column_dimensions[get_column_letter(j)].width = 9
        r = hr + 1
        groups = [("INDEX", [INDEX]), ("REGION", REGIONS), ("RATING", RATINGS),
                  ("COUNTRY", self.countries)]
        idx_row = None
        for gname, ents in groups:
            ents = [e for e in ents if any(self.p.get(d, e, M_STW) is not None for d in picks)]
            if not ents:
                continue
            r = band(ws, r, len(picks) + 1, gname)
            for e in ents:
                if e == INDEX:
                    idx_row = r
                L(ws.cell(row=r, column=1), e, bold=False)
                for j, d in enumerate(picks, start=2):
                    V(ws.cell(row=r, column=j), self.p.get(d, e, M_STW), fmt="0", color=C_HARD)
                r += 1
            r += 1
        ws.freeze_panes = "B5"
        if idx_row and len(picks) > 3:
            ch = LineChart()
            ch.title = "EMBIG Div spread (bp)"
            ch.height, ch.width = 8, 26
            ch.add_data(Reference(ws, min_col=2, max_col=len(picks) + 1,
                                  min_row=idx_row, max_row=idx_row), titles_from_data=False)
            ch.set_categories(Reference(ws, min_col=2, max_col=len(picks) + 1, min_row=hr, max_row=hr))
            ws.add_chart(ch, f"A{r + 1}")

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
        r = band(ws, r, 5, f"A. Spread responds to rates through beta ({BETAS['ust10y_bp']:.2f})")
        for i, h in enumerate(["UST 10y change", "implied spread (bp)", "spread change (bp)",
                               "12m total return %", "vs cash (pp)"], start=1):
            H(ws.cell(row=r, column=i), h)
        r += 1
        for u in SENSITIVITY_UST:
            sp = self.spread + BETAS["ust10y_bp"] * u
            tr = total_return(self.yld, self.IRD, self.D, u, self.spread, sp)
            L(ws.cell(row=r, column=1), f"{u:+d}bp", bold=(u == 0))
            V(ws.cell(row=r, column=2), sp, fmt="0", color=C_HARD)
            V(ws.cell(row=r, column=3), sp - self.spread, fmt="+0;-0")
            V(ws.cell(row=r, column=4), tr, fmt="0.00", bold=True,
              fill=C_GOOD if tr >= CASH_RATE else C_BAD)
            V(ws.cell(row=r, column=5), tr - CASH_RATE, fmt="+0.00;-0.00")
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
        self.history_tab()
        self.forecast()
        self.sensitivity()
        self.methodology()
        return self.wb


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
    for p in files:
        if p.name in (EXTENDED_CSV, EXTRAS_CSV) or not Panel.is_snapshot(p):
            continue
        res = panel.ingest_snapshot(p, known)
        if not res:
            continue
        d, nv, new = res
        print(f"Snapshot : {d}  {p.name}  ({nv} values, {new} new series)")
        added += 1
        archive.mkdir(exist_ok=True)
        dest = archive / f"snapshot_{d}.csv"
        if p.resolve() != dest.resolve() and not dest.exists():
            dest.write_bytes(p.read_bytes())

    if not panel.data:
        print("ERROR: no data at all.", file=sys.stderr)
        return 1

    dates = panel.dates()
    print(f"\nPanel    : {len(dates):,} dates, {dates[0]} to {dates[-1]}, "
          f"{len(panel.entities())} entities")

    d1 = dates[-1]
    d0 = year_start(panel, dt(d1).year) or dates[0]
    at = attribute(panel, INDEX, d0, d1) or {}
    D, dsrc = panel.duration(d1, INDEX)
    print(f"\nYTD {dt(d1).year} ({d0} -> {d1}), duration {D:.2f} ({dsrc})")
    print(f"   total return              {at.get('total', 0):+7.2f}%")
    print(f"   spread move               {at.get('spread_bp', 0):+7.0f}bp")
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
    Dashboard(panel).build().save(out)
    print(f"Wrote    : {out.name}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
