#!/usr/bin/env python3
"""
embig_history.py — EMBI history builder and dashboard.

WHY THIS EXISTS
---------------
JPM ships the same 962 bonds under more than one weighting scheme. On
03-Sep-2026 the download was 'EMBI Global' (Bam Id EMBIG_*); on 08-Sep-2026 it
was 'EMBI Global Diversified' (Bam Id FC_EMBIG_*). The bond universe is
IDENTICAL - same 962 issues, 151 issuers, 69 countries, byte-identical column
headers. Only the weights differ, because Diversified caps the larger-debt
issuers and redistributes to the smaller ones:

    Mexico 11.93% -> 5.12%     Saudi 11.98% -> 5.19%     Nigeria 1.49% -> 2.64%
    IG 55.7% -> 47.1%          Non-IG 44.3% -> 52.9%     index STW 188 -> 173

None of that is a market move. Splicing the two into one series would book a
15bp index tightening, and a 53bp LatAm tightening, that never happened - and
then feed that fiction into the forecast vol.

So this script treats the INDEX FAMILY as a first-class dimension. Every
observation is stamped with the family it came from, series are never mixed
across families, and the workbook is built for one family at a time.

THE HISTORY STORE
-----------------
History lives in ONE tidy, append-only CSV - embig_history.csv - with columns:

    date, family, section, entity, metric, value

Not in the workbook. The workbook is a disposable view; delete it and rerun and
you lose nothing. Every download appends its rows; re-running the same file
changes nothing (dedupe is on date+family+entity+metric, newest file wins).

This shape is the whole point. When JPM next changes something, a new column
becomes new METRIC ROWS rather than a schema break, and a new index becomes a
new FAMILY rather than a corrupted series. Nothing you have already collected
is invalidated by a change you have not seen yet.

USAGE
-----
    python embig_history.py                 # scan ., ingest, rebuild workbook
    python embig_history.py --dir /path/to/folder
    python embig_history.py --family "EMBI Global Diversified"
    python embig_history.py --no-build      # ingest only
    python embig_history.py --list          # what is in the store

Drop each new JPM download into the folder and rerun. Raw files are archived to
snapshots_archive/<family>/snapshot_<date>.csv so the store can always be
rebuilt from scratch.
"""

from __future__ import annotations

import argparse
import csv
import math
import sys
from collections import defaultdict
from datetime import datetime
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

from openpyxl import Workbook
from openpyxl.chart import LineChart, Reference
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

# ===========================================================================
# 1. CONFIGURATION
# ===========================================================================

HISTORY_FILE = "embig_history.csv"
ARCHIVE_DIR = "snapshots_archive"
DEFAULT_OUTPUT = "EMBI_Dashboard.xlsx"
DRIVERS_FILE = "drivers.csv"

# The spread we analyse. Alberto's call: duration-weighted spread-to-worst over
# the Treasury curve. The Z-spread is kept alongside because it is what lets a
# stripped-spread proxy be calibrated (see SPREAD_ANCHORS).
SPREAD_METRIC = "Dur Wgt STW (Trsy)"
ZSPREAD_METRIC = "Dur Wgt Z- Spread to Wrst"
YIELD_METRIC = "Dur Wgt Yield Wrst"

# --- stripped-spread calibration, PER FAMILY -------------------------------
# No JPM file carries a stripped spread. STW is measured off the Treasury par
# curve, Z off the zero curve, and the published stripped spread sits between
# them:  spread* = STW + theta * (Z - STW).
#
# theta is SOLVED from the index row against a published reference spread on a
# known date. It is family-specific - an EMBI Global anchor says nothing about
# Diversified. Set the anchor to None to publish raw STW for that family.
SPREAD_ANCHORS: Dict[str, Optional[Tuple[str, float]]] = {
    # family: (anchor date YYYY-MM-DD, published stripped spread in bp)
    "EMBI Global": ("2026-09-03", 218.42),
    "EMBI Global Diversified": None,   # <-- set once you have a reference print
}

# --- forecast model --------------------------------------------------------
# Practitioner priors, NOT in-sample estimates. Documented on the Methodology
# tab. beta(UST) is unconditional and that is its known weakness: it cannot
# tell "rates up on growth" (spreads tighten) from "rates up on an inflation
# shock" (spreads widen).
FORECAST_BETAS = {
    "ust_10y_bp": 0.30,     # per bp of UST 10y  -> bp of spread
    "dxy_pct": 6.0,         # per 1% of DXY      -> bp of spread
    "vix_pts": 4.0,         # per VIX point      -> bp of spread
    "rating_notches": -50.0,  # per notch of upgrade drift -> bp of spread
}
FORECAST_PERCENTILES = [5, 25, 50, 75, 95]

# How many rows the Drivers sheet offers for pasting history. The regression
# reads the whole block and ignores blanks, so oversizing costs nothing.
DRIVER_ROWS = 160

# Minimum observations before a fitted beta is worth looking at. Below this the
# Regression tab still computes but flags itself as not usable.
MIN_OBS_FOR_REGRESSION = 24

# Rows whose Instrument is one of these are section separators, not data.
SECTION_LABELS = {
    "By Region": "region",
    "By Country": "country",
    "By Credit Bucket": "credit",
    "By Sub Credit Bucket": "subcredit",
    "By Sov/Quasi Sov/Corp": "structure",
}

# Every numeric column worth keeping. Anything not listed is ignored; anything
# new that JPM adds can simply be appended here.
NUMERIC_METRICS = [
    "Index Level", "Mkt Cap", "Mkt Cap %",
    "Daily Change (%)", "MTD Change (%)", "YTD Change (%)",
    "UST Return Daily Change (%)", "UST Return MTD Change (%)", "UST Return YTD Change (%)",
    "Spread Return Daily Change (%)", "Spread Return MTD Change (%)", "Spread Return YTD Change (%)",
    "Price Return YTD Change (%)", "Coupon Return YTD Change (%)",
    "No. of Issues", "No. of Issuer",
    "Coupon Amt", "Spread Duration", "IR Duration to Worst",
    "Weighted Avg Avg Life Wrst",
    YIELD_METRIC, SPREAD_METRIC, ZSPREAD_METRIC,
]
TEXT_METRICS = ["Average S&P Rating", "Average Moody Rating", "Average Fitch Rating"]

REGION_ORDER = ["Latin", "Europe", "Africa", "Middle East", "Asia"]
RATING_ORDER = ["A", "BBB", "BB", "B", "C", "NR"]
CREDIT_ORDER = ["Investment Grade", "Non Investment Grade"]

# ===========================================================================
# 2. STYLING
# ===========================================================================

FONT_NAME = "Arial"
C_HDR_BG, C_HDR_FG = "1F4E78", "FFFFFF"
C_BAND, C_INDEX = "D9E1F2", "C6E0B4"
C_INPUT, C_BORDER = "FFFF00", "BFBFBF"
C_HARD, C_FORMULA, C_NOTE = "0000FF", "000000", "7F7F7F"
BORDER = Border(*[Side(style="thin", color=C_BORDER)] * 4)


def _font(cell, **kw):
    cell.font = Font(name=FONT_NAME, size=kw.get("size", 10), bold=kw.get("bold", False),
                     italic=kw.get("italic", False), color=kw.get("color", C_FORMULA))


def _hdr(cell, txt):
    cell.value = txt
    cell.fill = PatternFill("solid", fgColor=C_HDR_BG)
    cell.font = Font(name=FONT_NAME, size=10, bold=True, color=C_HDR_FG)
    cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    cell.border = BORDER


def _val(cell, value, fmt="0.00", color=C_FORMULA, bold=False):
    cell.value = value
    cell.number_format = fmt
    cell.border = BORDER
    _font(cell, color=color, bold=bold)


def _label(cell, txt, fill=None, bold=True):
    cell.value = txt
    cell.border = BORDER
    if fill:
        cell.fill = PatternFill("solid", fgColor=fill)
    _font(cell, bold=bold)


# ===========================================================================
# 3. PARSING
# ===========================================================================

def _parse_date(raw: str) -> datetime:
    raw = (raw or "").strip().strip('"')
    for fmt in ("%d-%b-%Y", "%Y-%m-%d", "%m/%d/%Y", "%d/%m/%Y", "%b %d, %Y", "%d-%b-%y"):
        try:
            return datetime.strptime(raw, fmt)
        except ValueError:
            continue
    raise ValueError(f"unparseable date: {raw!r}")


def _num(raw: Any) -> Optional[float]:
    if raw is None:
        return None
    s = str(raw).strip().replace(",", "").replace("%", "")
    if not s or s in {"-", "--", "n/a", "N/A", "NA", "#N/A"}:
        return None
    if s.startswith("(") and s.endswith(")"):
        s = "-" + s[1:-1]
    try:
        v = float(s)
    except ValueError:
        return None
    return None if math.isnan(v) or math.isinf(v) else v


def is_jpm_snapshot(path: Path) -> bool:
    """A JPM sub-index dump is identified by its FIRST COLUMN being 'Bam Id'.
    Filenames are worthless here - JPM changes them every release."""
    try:
        with path.open("r", newline="", encoding="utf-8-sig", errors="replace") as f:
            for row in csv.reader(f):
                if any((c or "").strip() for c in row):
                    return (row[0] or "").strip().strip('"').lower() in {"bam id", "bamid", "bam_id"}
    except OSError:
        return False
    return False


def detect_family(index_instrument: str, bam_id: str) -> str:
    """Which index this file is. Read from the index row's own label first -
    that is JPM telling us directly. Bam Id prefix is the fallback."""
    t = (index_instrument or "").strip()
    if t:
        return t
    return "EMBI Global Diversified" if (bam_id or "").upper().startswith("FC_") else "EMBI Global"


def parse_snapshot(path: Path) -> Optional[Dict[str, Any]]:
    """Return {date, family, rows: [(section, entity, metric, value)]}."""
    with path.open("r", newline="", encoding="utf-8-sig", errors="replace") as f:
        raw = [{(k or "").strip(): v for k, v in r.items()} for r in csv.DictReader(f)]
    if not raw:
        return None

    snap_date: Optional[datetime] = None
    family: Optional[str] = None
    index_bam = ""
    section = "index"
    out: List[Tuple[str, str, str, Any]] = []

    for r in raw:
        bam = (r.get("Bam Id") or "").strip()
        inst = (r.get("Instrument") or "").strip()

        # trailing disclaimer block has no Instrument and a junk Bam Id
        if bam.lower().startswith(("disclaimer", "http")):
            break
        if inst in SECTION_LABELS:
            section = SECTION_LABELS[inst]
            continue
        if not inst:
            continue
        if snap_date is None:
            try:
                snap_date = _parse_date(r.get("Date", ""))
            except ValueError:
                pass
        if family is None:          # first data row IS the index row
            family = detect_family(inst, bam)
            index_bam = bam
            entity = "INDEX"
        else:
            entity = inst
        # the index row sits before any "By ..." marker
        sec = "index" if entity == "INDEX" else section

        for m in NUMERIC_METRICS:
            v = _num(r.get(m))
            if v is not None:
                out.append((sec, entity, m, v))
        for m in TEXT_METRICS:
            v = (r.get(m) or "").strip()
            if v and v.upper() not in {"NR", ""}:
                out.append((sec, entity, m, v))

    if snap_date is None or family is None:
        return None
    return {"date": snap_date, "family": family, "bam": index_bam, "rows": out}


# ===========================================================================
# 4. THE HISTORY STORE
# ===========================================================================

class History:
    """Append-only tidy store: (date, family, section, entity, metric) -> value."""

    FIELDS = ["date", "family", "section", "entity", "metric", "value"]

    def __init__(self) -> None:
        self.data: Dict[Tuple[str, str, str, str, str], Any] = {}

    # ---- persistence ----
    def load(self, path: Path) -> None:
        if not path.exists():
            return
        with path.open("r", newline="", encoding="utf-8") as f:
            for r in csv.DictReader(f):
                key = (r["date"], r["family"], r["section"], r["entity"], r["metric"])
                self.data[key] = r["value"]

    def save(self, path: Path) -> None:
        with path.open("w", newline="", encoding="utf-8") as f:
            w = csv.writer(f)
            w.writerow(self.FIELDS)
            for (d, fam, sec, ent, met), v in sorted(self.data.items()):
                w.writerow([d, fam, sec, ent, met, v])

    # ---- ingestion ----
    def ingest(self, snap: Dict[str, Any]) -> Tuple[int, int]:
        d = snap["date"].strftime("%Y-%m-%d")
        fam = snap["family"]
        added = updated = 0
        for sec, ent, met, val in snap["rows"]:
            key = (d, fam, sec, ent, met)
            if key in self.data:
                if str(self.data[key]) != str(val):
                    updated += 1
            else:
                added += 1
            self.data[key] = val
        return added, updated

    # ---- reading ----
    def families(self) -> List[str]:
        return sorted({k[1] for k in self.data})

    def dates(self, family: str) -> List[str]:
        return sorted({k[0] for k in self.data if k[1] == family})

    def entities(self, family: str, section: str) -> List[str]:
        return sorted({k[3] for k in self.data if k[1] == family and k[2] == section})

    def get(self, family: str, entity: str, metric: str, date: str) -> Optional[Any]:
        for sec in ("index", "region", "country", "credit", "subcredit", "structure"):
            v = self.data.get((date, family, sec, entity, metric))
            if v is not None:
                return v
        return None

    def num(self, family: str, entity: str, metric: str, date: str) -> Optional[float]:
        return _num(self.get(family, entity, metric, date))

    def series(self, family: str, entity: str, metric: str) -> List[Tuple[str, float]]:
        out = []
        for d in self.dates(family):
            v = self.num(family, entity, metric, d)
            if v is not None:
                out.append((d, v))
        return out

    def latest(self, family: str) -> Optional[str]:
        ds = self.dates(family)
        return ds[-1] if ds else None


# ===========================================================================
# 4b. DRIVERS AND THE REGRESSION (fitted here in Python, not in Excel)
# ===========================================================================

# WHICH TREASURY DRIVES THE SPREAD?
# -----------------------------------------------------------------------
# A duration of ~6.2 tempts you towards the 5y, but duration is a SENSITIVITY,
# not a location on the curve. The index's cash flows sit at 10.2y average
# life; duration is shorter only because coupons pay along the way.
#
# The file settles it. Backing the Treasury anchor out of each country's own
# numbers (implied UST = yield - STW) and regressing on maturity across 66
# countries:
#
#     vs average life        slope +0.0489/yr   R2 0.881   <- best fit
#     vs spread duration     slope +0.0985/yr   R2 0.842
#
# The index anchor is 4.916%, which on that fitted curve is an 11.2y point.
# The 5y point is 4.615% — 30bp below where the index is actually struck.
# So the 10y is the right level factor, and the 5y would be measuring a rate
# the index is not exposed to.
#
# The exception is HEDGING: a 6.2-duration book against a ~8-duration 10y note
# is a 0.78 hedge ratio, and a 5y/10y blend is a legitimate way to build it.
# That is a different question from which rate explains spread changes.
#
# Set RATE_COL to "ust5y" to test it anyway, or use --rate compare to fit both
# and let your own data decide. Note the 0.30 prior was calibrated on 10y
# moves; it does not carry over to a 5y specification.
# A curve TWIST hits EM differently from a parallel shift - a bear flattening
# is a policy shock, a bear steepening is usually growth or supply - so the
# default specification carries level AND slope. Whether slope earns its
# degree of freedom is settled by adjusted R-squared, not by assertion:
# run --spec compare.
RATE_SPECS: Dict[str, List[str]] = {
    "level":       ["ust10y", "dxy", "vix"],
    "level_slope": ["ust10y", "slope2s10s", "dxy", "vix"],
    "front":       ["ust5y", "dxy", "vix"],
}
SPEC = "level_slope"

DRIVER_LABELS = {
    "ust10y": "UST 10y level (per bp)", "ust5y": "UST 5y level (per bp)",
    "slope2s10s": "2s10s slope (per bp)", "dxy": "DXY (per 1%)", "vix": "VIX (per point)",
}
DRIVER_PRIORS = {"ust10y": "ust_10y_bp", "ust5y": "ust_10y_bp",
                 "slope2s10s": None, "dxy": "dxy_pct", "vix": "vix_pts"}
# Raw columns drivers.csv may carry. slope2s10s is derived, not read.
RAW_RATE_COLS = ["ust2y", "ust5y", "ust10y"]
DRIVER_COLS = RATE_SPECS[SPEC]


def load_drivers(path: Path) -> Dict[str, Dict[str, Optional[float]]]:
    """drivers.csv — one row per observation:
           date,ust10y,dxy,vix[,regime]
       regime is optional and MUST be exogenous to the model (see fit_ols)."""
    out: Dict[str, Dict[str, Optional[float]]] = {}
    if not path.exists():
        return out
    with path.open("r", newline="", encoding="utf-8-sig") as f:
        for r in csv.DictReader(f):
            r = {(k or "").strip().lower(): v for k, v in r.items()}
            raw = (r.get("date") or "").strip()
            if not raw or raw.startswith("#"):
                continue
            try:
                d = _parse_date(raw).strftime("%Y-%m-%d")
            except ValueError:
                continue
            out[d] = {c: _num(r.get(c)) for c in set(RAW_RATE_COLS) | {"dxy", "vix"}}
            out[d]["regime"] = _num(r.get("regime"))
    return out


def write_drivers_template(path: Path) -> None:
    """A starter drivers.csv. Comment lines begin with '#' and are ignored."""
    lines = [
        "# One row per month, oldest first. Yields in PERCENT (4.57, not 0.0457).",
        "# ust10y is required. ust2y feeds the 2s10s slope factor used by the default",
        "#   'level_slope' specification. ust5y is only needed for --spec front.",
        "# regime: 1 = the rate move came with risk aversion (inflation/policy shock),",
        "#         0 = it came with growth. Leave blank if you have no view - the",
        "#         conditional block is then skipped rather than guessed.",
        "# Do NOT derive regime from VIX: VIX is already a regressor, and conditioning",
        "#   on it manufactures a difference between the two betas even when none exists.",
        "date,ust2y,ust5y,ust10y,dxy,vix,regime",
        "2026-01-31,,,,,,",
    ]
    path.write_text("\n".join(lines) + "\n", encoding="utf-8")


def _solve(A: List[List[float]], b: List[float]) -> Optional[List[float]]:
    """Gauss-Jordan with partial pivoting. Returns None if singular."""
    n = len(A)
    M = [row[:] + [b[i]] for i, row in enumerate(A)]
    for c in range(n):
        p = max(range(c, n), key=lambda r: abs(M[r][c]))
        if abs(M[p][c]) < 1e-12:
            return None
        M[c], M[p] = M[p], M[c]
        pv = M[c][c]
        M[c] = [x / pv for x in M[c]]
        for r in range(n):
            if r != c and M[r][c] != 0.0:
                fac = M[r][c]
                M[r] = [a - fac * bb for a, bb in zip(M[r], M[c])]
    return [M[i][n] for i in range(n)]


def _inverse(A: List[List[float]]) -> Optional[List[List[float]]]:
    n = len(A)
    cols = []
    for i in range(n):
        e = [1.0 if j == i else 0.0 for j in range(n)]
        c = _solve(A, e)
        if c is None:
            return None
        cols.append(c)
    return [[cols[j][i] for j in range(n)] for i in range(n)]


def ols(y: List[float], X: List[List[float]]) -> Optional[Dict[str, Any]]:
    """Plain OLS with an intercept. X is a list of rows (no constant column)."""
    n, k = len(y), len(X[0]) if X else 0
    if n < k + 2:
        return None
    Z = [[1.0] + row for row in X]
    p = k + 1
    XtX = [[sum(Z[i][a] * Z[i][b] for i in range(n)) for b in range(p)] for a in range(p)]
    Xty = [sum(Z[i][a] * y[i] for i in range(n)) for a in range(p)]
    beta = _solve(XtX, Xty)
    if beta is None:
        return None
    fitted = [sum(beta[a] * Z[i][a] for a in range(p)) for i in range(n)]
    resid = [y[i] - fitted[i] for i in range(n)]
    dof = n - p
    if dof <= 0:
        return None
    ssr = sum(r * r for r in resid)
    ybar = sum(y) / n
    sst = sum((v - ybar) ** 2 for v in y)
    s2 = ssr / dof
    inv = _inverse(XtX)
    se = [math.sqrt(max(s2 * inv[a][a], 0.0)) for a in range(p)] if inv else [float("nan")] * p
    r2 = (1 - ssr / sst) if sst > 0 else float("nan")
    adj = 1 - (1 - r2) * (n - 1) / dof if sst > 0 else float("nan")
    dw = (sum((resid[i] - resid[i - 1]) ** 2 for i in range(1, n)) / ssr) if ssr > 0 else float("nan")
    # Variance inflation: regress each driver on the others. >5 means the
    # coefficients are being split between collinear factors and the individual
    # t-stats stop meaning much - the usual fate of level-and-slope together.
    vif = []
    for j in range(k):
        yj = [row[j] for row in X]
        Xj = [[row[m] for m in range(k) if m != j] for row in X]
        if k == 1 or not Xj or not Xj[0]:
            vif.append(1.0); continue
        fj = ols(yj, Xj)
        vif.append(1.0 / (1.0 - fj["r2"]) if fj and fj["r2"] < 0.9999 else float("inf"))
    return {
        "n": n, "beta": beta[1:], "se": se[1:], "intercept": beta[0], "intercept_se": se[0],
        "t": [(beta[a] / se[a] if se[a] else float("nan")) for a in range(1, p)],
        "r2": r2, "adj_r2": adj, "dw": dw, "vif": vif,
        "se_reg": math.sqrt(s2),
        "sigma_y": math.sqrt(sum((v - ybar) ** 2 for v in y) / n),
    }


def build_panel(hist: "History", family: str,
                drivers: Dict[str, Dict[str, Optional[float]]]) -> List[Dict[str, Any]]:
    """Monthly CHANGES, aligned on date. Levels would give a flattering R-squared
    on two near-random-walks and mean nothing."""
    dates = [d for d in hist.dates(family)
             if hist.num(family, "INDEX", SPREAD_METRIC, d) is not None]
    rows = []
    for i in range(1, len(dates)):
        d0, d1 = dates[i - 1], dates[i]
        if d0 not in drivers or d1 not in drivers:
            continue
        a, b = drivers[d0], drivers[d1]
        if a.get("dxy") in (None, 0) or b.get("dxy") is None \
           or a.get("vix") is None or b.get("vix") is None:
            continue
        rates: Dict[str, float] = {}
        for rc in RAW_RATE_COLS:
            if a.get(rc) is not None and b.get(rc) is not None:
                rates[rc] = (b[rc] - a[rc]) * 100.0        # to bp
        if a.get("ust10y") is not None and a.get("ust2y") is not None \
           and b.get("ust10y") is not None and b.get("ust2y") is not None:
            rates["slope2s10s"] = ((b["ust10y"] - b["ust2y"]) - (a["ust10y"] - a["ust2y"])) * 100.0
        s0 = hist.num(family, "INDEX", SPREAD_METRIC, d0)
        s1 = hist.num(family, "INDEX", SPREAD_METRIC, d1)
        rows.append({
            "date": d1,
            "d_spread": s1 - s0,
            "dxy": (b["dxy"] / a["dxy"] - 1.0) * 100.0,       # to %
            "vix": b["vix"] - a["vix"],
            "regime": b.get("regime"),
            **rates,
        })
    return rows


def usable(panel: List[Dict[str, Any]], cols: List[str]) -> List[Dict[str, Any]]:
    return [r for r in panel if all(r.get(c) is not None for c in cols)]


def fit_regression(panel: List[Dict[str, Any]],
                   cols: Optional[List[str]] = None) -> Dict[str, Any]:
    cols = cols or DRIVER_COLS
    sub = usable(panel, cols)
    out: Dict[str, Any] = {"panel_n": len(sub), "cols": cols, "fit": None, "conditional": {}}
    if len(sub) < 5:
        return out
    y = [r["d_spread"] for r in sub]
    X = [[r[c] for c in cols] for r in sub]
    out["fit"] = ols(y, X)
    panel = sub

    # Conditional UST beta, split on the EXOGENOUS regime flag the user supplies.
    # Never split on VIX: it is a regressor, so conditioning on it produces a
    # spurious difference between the two betas even when none exists.
    for label, want in (("growth", 0.0), ("shock", 1.0)):
        sub = [r for r in panel if r.get("regime") is not None and float(r["regime"]) == want]
        if len(sub) >= 4:
            f = ols([r["d_spread"] for r in sub], [[r["ust10y"]] for r in sub])
            if f:
                out["conditional"][label] = {"n": f["n"], "beta": f["beta"][0],
                                             "se": f["se"][0], "t": f["t"][0]}
        elif sub:
            out["conditional"][label] = {"n": len(sub), "beta": None, "se": None, "t": None}
    return out


def compare_specs(panel: List[Dict[str, Any]]) -> Dict[str, Optional[Dict[str, Any]]]:
    """Fit every rate specification on the same panel. Compare on ADJUSTED
    R-squared: plain R-squared can only rise when a factor is added, so it
    always votes for the bigger model."""
    out: Dict[str, Optional[Dict[str, Any]]] = {}
    for name, cols in RATE_SPECS.items():
        sub = usable(panel, cols)
        if len(sub) < 5:
            out[name] = None
            continue
        f = ols([r["d_spread"] for r in sub], [[r[c] for c in cols] for r in sub])
        if f:
            f["cols"] = cols
        out[name] = f
    return out


def chosen_betas(res: Dict[str, Any]) -> Dict[str, Tuple[float, str]]:
    """Fitted beta only when the sample is big enough AND the coefficient is
    significant; otherwise the practitioner prior. Returns {key: (value, why)}."""
    out = {}
    fit = res.get("fit")
    cols = res.get("cols") or DRIVER_COLS
    for i, c in enumerate(cols):
        pk = DRIVER_PRIORS.get(c)
        prior = FORECAST_BETAS[pk] if pk else 0.0
        if fit and fit["n"] >= MIN_OBS_FOR_REGRESSION and abs(fit["t"][i]) >= 2.0:
            out[c] = (fit["beta"][i], f"fitted (n={fit['n']}, t={fit['t'][i]:+.2f})")
        elif fit:
            why = (f"prior — n={fit['n']} < {MIN_OBS_FOR_REGRESSION}"
                   if fit["n"] < MIN_OBS_FOR_REGRESSION
                   else f"prior — |t|={abs(fit['t'][i]):.2f} < 2")
            out[c] = (prior, why)
        else:
            out[c] = (prior, "prior — not enough data to fit")
    return out


# ===========================================================================
# 5. STRIPPED-SPREAD BASIS
# ===========================================================================

def solve_theta(hist: History, family: str) -> Optional[float]:
    """theta = (anchor - STW) / (Z - STW), solved ONCE on the anchor date.

    Never re-solve per date: the anchor is one observed stripped spread on one
    day. Re-solving against every snapshot would force every date to print the
    anchor and flatten the spread series into a constant. theta is a basis
    conversion; what moves through time is each date's own (Z - STW)."""
    anchor = SPREAD_ANCHORS.get(family)
    if not anchor:
        return None
    date, target = anchor
    s = hist.num(family, "INDEX", SPREAD_METRIC, date)
    z = hist.num(family, "INDEX", ZSPREAD_METRIC, date)
    if s is None or z is None or abs(z - s) < 1.0:
        return None
    return (target - s) / (z - s)


def stripped(hist: History, family: str, entity: str, date: str,
             theta: Optional[float]) -> Optional[float]:
    """STW re-based onto the stripped basis. Where Z < STW the two measures
    have inverted - that only happens on defaulted paper whose yield-to-worst
    is an artefact (Venezuela, Lebanon) - so those stay on raw STW."""
    s = hist.num(family, entity, SPREAD_METRIC, date)
    if s is None or theta is None:
        return s
    z = hist.num(family, entity, ZSPREAD_METRIC, date)
    if z is None or z < s:
        return s
    return s + theta * (z - s)


# ===========================================================================
# 6. WORKBOOK
# ===========================================================================

class Dashboard:
    def __init__(self, hist: History, family: str,
                 drivers: Optional[Dict[str, Dict[str, Optional[float]]]] = None) -> None:
        self.h = hist
        self.fam = family
        self.dates = hist.dates(family)
        self.d = self.dates[-1] if self.dates else None
        self.theta = solve_theta(hist, family)
        self.drv = drivers or {}
        self.panel = build_panel(hist, family, self.drv)
        self.reg = fit_regression(self.panel)
        self.wb = Workbook()
        self.wb.remove(self.wb.active)

    # ---------------- cover ----------------
    def cover(self):
        ws = self.wb.create_sheet("Cover")
        ws.sheet_view.showGridLines = False
        ws["A1"] = f"{self.fam} — spreads, performance and forecast"
        _font(ws["A1"], size=16, bold=True, color=C_HDR_BG)
        ws.column_dimensions["A"].width = 42
        ws.column_dimensions["B"].width = 60

        anchor = SPREAD_ANCHORS.get(self.fam)
        rows = [
            ("Index family", self.fam),
            ("Latest observation", self.d or "—"),
            ("Observations in store", len(self.dates)),
            ("History spans", f"{self.dates[0]} to {self.dates[-1]}" if self.dates else "—"),
            ("Countries", len(self.h.entities(self.fam, "country"))),
            ("Spread metric", SPREAD_METRIC),
            ("Stripped basis", (f"theta {self.theta:.4f} solved on {anchor[0]} "
                                f"against {anchor[1]}bp") if self.theta else
                               "not calibrated — showing raw STW"),
            ("Generated", datetime.now().strftime("%Y-%m-%d %H:%M")),
        ]
        r = 3
        for k, v in rows:
            _label(ws.cell(row=r, column=1), k, fill=C_BAND)
            _val(ws.cell(row=r, column=2), v, fmt="General", color=C_HARD)
            r += 1

        r += 1
        ws.cell(row=r, column=1, value="Index families held in this store — never spliced:")
        _font(ws.cell(row=r, column=1), bold=True, color=C_HDR_BG)
        r += 1
        for fam in self.h.families():
            ds = self.h.dates(fam)
            mark = "  <- this workbook" if fam == self.fam else ""
            ws.cell(row=r, column=1, value=fam)
            ws.cell(row=r, column=2, value=f"{len(ds)} obs, {ds[0]} to {ds[-1]}{mark}")
            _font(ws.cell(row=r, column=1)); _font(ws.cell(row=r, column=2), color=C_NOTE)
            r += 1

        r += 1
        for line in [
            "EMBI Global and EMBI Global Diversified hold the SAME 962 bonds under different",
            "weights. Diversified caps the larger-debt issuers, so Mexico falls from 11.9% to",
            "5.1%, non-IG rises from 44% to 53%, and the index spread prints 173 instead of 188.",
            "None of that is a market move. The two families are stored and charted separately;",
            "joining them would book a 15bp tightening that never happened.",
        ]:
            ws.cell(row=r, column=1, value=line)
            _font(ws.cell(row=r, column=1), italic=True, color=C_NOTE)
            r += 1

    # ---------------- country spreads ----------------
    def country_spreads(self):
        ws = self.wb.create_sheet("Country_Spreads")
        ws.sheet_view.showGridLines = False
        ws["A1"] = f"Country spreads — {self.fam}, {self.d}"
        _font(ws["A1"], size=14, bold=True, color=C_HDR_BG)
        ws["A2"] = ("Spread is Dur Wgt STW (Trsy). 'Stripped proxy' re-bases it onto the published "
                    "stripped basis via theta; blank when the family has no anchor.")
        _font(ws["A2"], italic=True, color=C_NOTE)

        cols = [("Country", 26), ("S&P", 8), ("Moody's", 9), ("Weight %", 10),
                ("STW (bp)", 10), ("Z-spread (bp)", 12), ("Stripped proxy", 13),
                ("Spread dur", 11), ("Yield %", 10), ("Avg life", 10),
                ("YTD TR %", 10), ("of which spread", 14), ("of which UST", 12)]
        hr = 4
        for i, (h, w) in enumerate(cols, start=1):
            _hdr(ws.cell(row=hr, column=i), h)
            ws.column_dimensions[get_column_letter(i)].width = w

        countries = self.h.entities(self.fam, "country")
        rows = []
        for c in countries:
            rows.append((c, self.h.num(self.fam, c, "Mkt Cap %", self.d) or 0.0))
        rows.sort(key=lambda t: -t[1])

        r = hr + 1
        for c, wgt in rows:
            g = lambda m: self.h.num(self.fam, c, m, self.d)
            _label(ws.cell(row=r, column=1), c, bold=False)
            for col, m in ((2, "Average S&P Rating"), (3, "Average Moody Rating")):
                cell = ws.cell(row=r, column=col, value=self.h.get(self.fam, c, m, self.d) or "")
                cell.border = BORDER; cell.alignment = Alignment(horizontal="center"); _font(cell)
            _val(ws.cell(row=r, column=4), wgt, fmt="0.00", color=C_HARD)
            _val(ws.cell(row=r, column=5), g(SPREAD_METRIC), fmt="0", color=C_HARD)
            _val(ws.cell(row=r, column=6), g(ZSPREAD_METRIC), fmt="0", color=C_HARD)
            sp = stripped(self.h, self.fam, c, self.d, self.theta)
            _val(ws.cell(row=r, column=7), sp if self.theta else "", fmt="0.0")
            _val(ws.cell(row=r, column=8), g("Spread Duration"), fmt="0.00", color=C_HARD)
            _val(ws.cell(row=r, column=9), g(YIELD_METRIC), fmt="0.00", color=C_HARD)
            _val(ws.cell(row=r, column=10), g("Weighted Avg Avg Life Wrst"), fmt="0.0", color=C_HARD)
            _val(ws.cell(row=r, column=11), g("YTD Change (%)"), fmt="+0.00;-0.00", color=C_HARD)
            _val(ws.cell(row=r, column=12), g("Spread Return YTD Change (%)"), fmt="+0.00;-0.00", color=C_HARD)
            _val(ws.cell(row=r, column=13), g("UST Return YTD Change (%)"), fmt="+0.00;-0.00", color=C_HARD)
            r += 1
        ws.freeze_panes = "B5"

    # ---------------- performance blocks ----------------
    def _perf_block(self, ws, row, title, section, order):
        _label(ws.cell(row=row, column=1), title, fill=C_BAND)
        for c in range(2, 12):
            ws.cell(row=row, column=c).fill = PatternFill("solid", fgColor=C_BAND)
            ws.cell(row=row, column=c).border = BORDER
        row += 1
        present = self.h.entities(self.fam, section)
        ordered = [e for e in order if e in present] + [e for e in present if e not in order]
        for e in ordered:
            g = lambda m: self.h.num(self.fam, e, m, self.d)
            _label(ws.cell(row=row, column=1), self.fam if e == "INDEX" else e, bold=False)
            _val(ws.cell(row=row, column=2), g("Mkt Cap %"), fmt="0.00", color=C_HARD)
            _val(ws.cell(row=row, column=3), g(SPREAD_METRIC), fmt="0", color=C_HARD)
            sp = stripped(self.h, self.fam, e, self.d, self.theta)
            _val(ws.cell(row=row, column=4), sp if self.theta else "", fmt="0.0")
            _val(ws.cell(row=row, column=5), g(YIELD_METRIC), fmt="0.00", color=C_HARD)
            _val(ws.cell(row=row, column=6), g("Spread Duration"), fmt="0.00", color=C_HARD)
            _val(ws.cell(row=row, column=7), g("IR Duration to Worst"), fmt="0.00", color=C_HARD)
            _val(ws.cell(row=row, column=8), g("YTD Change (%)"), fmt="+0.00;-0.00", color=C_HARD, bold=True)
            _val(ws.cell(row=row, column=9), g("Spread Return YTD Change (%)"), fmt="+0.00;-0.00", color=C_HARD)
            _val(ws.cell(row=row, column=10), g("UST Return YTD Change (%)"), fmt="+0.00;-0.00", color=C_HARD)
            # Spread share is meaningless when total return is ~0 — a 0.05%
            # denominator turns a normal attribution into a 200% headline.
            tot = g("YTD Change (%)"); spr = g("Spread Return YTD Change (%)")
            share = (spr / tot) if (tot is not None and abs(tot) >= 0.25 and spr is not None) else ""
            _val(ws.cell(row=row, column=11), share, fmt="0%")
            if share == "" and tot is not None:
                ws.cell(row=row, column=11, value="n/m").alignment = Alignment(horizontal="right")
                _font(ws.cell(row=row, column=11), color=C_NOTE, italic=True)
                ws.cell(row=row, column=11).border = BORDER
            row += 1
        return row + 1

    def performance(self):
        ws = self.wb.create_sheet("Performance_YTD")
        ws.sheet_view.showGridLines = False
        ws["A1"] = f"Year-to-date performance — {self.fam}, {self.d}"
        _font(ws["A1"], size=14, bold=True, color=C_HDR_BG)
        ws["A2"] = ("Total return split into its spread and Treasury legs, straight from JPM's own "
                    "attribution columns. 'Spread share' is the fraction of total return the spread "
                    "leg accounts for — above 100% means rates worked against you.")
        _font(ws["A2"], italic=True, color=C_NOTE)

        cols = [("", 26), ("Weight %", 10), ("STW (bp)", 10), ("Stripped", 10), ("Yield %", 9),
                ("Spread dur", 11), ("IR dur", 9), ("YTD TR %", 10),
                ("of which: spread", 15), ("of which: UST", 13), ("Spread share", 12)]
        hr = 4
        for i, (h, w) in enumerate(cols, start=1):
            _hdr(ws.cell(row=hr, column=i), h)
            ws.column_dimensions[get_column_letter(i)].width = w

        r = hr + 1
        idx = self.h.entities(self.fam, "index")
        r = self._perf_block(ws, r, "INDEX", "index", idx)
        r = self._perf_block(ws, r, "BY REGION", "region", REGION_ORDER)
        r = self._perf_block(ws, r, "BY CREDIT", "credit", CREDIT_ORDER)
        r = self._perf_block(ws, r, "BY RATING", "subcredit", RATING_ORDER)
        r = self._perf_block(ws, r, "BY STRUCTURE", "structure", ["Sovereign", "Quasi"])
        ws.freeze_panes = "B5"

    # ---------------- spread history ----------------
    def spread_history(self):
        ws = self.wb.create_sheet("Spread_History")
        ws.sheet_view.showGridLines = False
        ws["A1"] = f"Spread history — {self.fam}"
        _font(ws["A1"], size=14, bold=True, color=C_HDR_BG)
        ws["A2"] = ("One column per download. This grows every time you drop a new JPM file in the "
                    "folder and rerun — the store is what accumulates, not this sheet.")
        _font(ws["A2"], italic=True, color=C_NOTE)

        hr = 4
        _hdr(ws.cell(row=hr, column=1), "Entity")
        ws.column_dimensions["A"].width = 28
        for j, d in enumerate(self.dates, start=2):
            _hdr(ws.cell(row=hr, column=j), d)
            ws.column_dimensions[get_column_letter(j)].width = 12

        r = hr + 1
        groups = [("INDEX", ["INDEX"]),
                  ("REGION", [e for e in REGION_ORDER if e in self.h.entities(self.fam, "region")]),
                  ("RATING", [e for e in RATING_ORDER if e in self.h.entities(self.fam, "subcredit")]),
                  ("COUNTRY", self.h.entities(self.fam, "country"))]
        for gname, ents in groups:
            if not ents:
                continue
            _label(ws.cell(row=r, column=1), gname, fill=C_BAND)
            for j in range(2, len(self.dates) + 2):
                ws.cell(row=r, column=j).fill = PatternFill("solid", fgColor=C_BAND)
                ws.cell(row=r, column=j).border = BORDER
            r += 1
            for e in ents:
                _label(ws.cell(row=r, column=1), e, bold=False)
                for j, d in enumerate(self.dates, start=2):
                    v = stripped(self.h, self.fam, e, d, self.theta) if self.theta \
                        else self.h.num(self.fam, e, SPREAD_METRIC, d)
                    _val(ws.cell(row=r, column=j), v if v is not None else "", fmt="0", color=C_HARD)
                r += 1
            r += 1
        ws.freeze_panes = "B5"

        if len(self.dates) >= 2:
            ch = LineChart()
            ch.title = f"{self.fam} index spread (bp)"
            ch.height, ch.width = 8, 18
            data = Reference(ws, min_col=2, max_col=len(self.dates) + 1, min_row=hr + 2, max_row=hr + 2)
            cats = Reference(ws, min_col=2, max_col=len(self.dates) + 1, min_row=hr, max_row=hr)
            ch.add_data(data, titles_from_data=False)
            ch.set_categories(cats)
            ws.add_chart(ch, f"A{r + 2}")

    # ---------------- drivers (the regression inputs, as data) ----------------
    def drivers(self):
        ws = self.wb.create_sheet("Drivers")
        ws.sheet_view.showGridLines = False
        ws["A1"] = f"Regression inputs — {self.fam}"
        _font(ws["A1"], size=14, bold=True, color=C_HDR_BG)
        ws["A2"] = (f"Source of truth is {DRIVERS_FILE} in the project folder — edit that and rerun. "
                    "This tab is the audit trail, not an input form.")
        _font(ws["A2"], italic=True, color=C_NOTE)
        ws["A3"] = ("Spread comes from the history store; UST/DXY/VIX come from drivers.csv. Only "
                    "dates present in BOTH are used, which is why the count here can be lower than "
                    "either source on its own.")
        _font(ws["A3"], italic=True, color=C_NOTE)

        cols = [("Date", 14), ("Spread (bp)", 13), ("UST 10y (%)", 12), ("DXY", 11), ("VIX", 10),
                ("Regime", 10), ("d spread (bp)", 14), ("d UST (bp)", 12), ("d DXY (%)", 12),
                ("d VIX (pts)", 12), ("in regression", 14)]
        hr = 5
        for i, (h, w) in enumerate(cols, start=1):
            _hdr(ws.cell(row=hr, column=i), h)
            ws.column_dimensions[get_column_letter(i)].width = w

        used = {r["date"] for r in self.panel}
        by_date = {r["date"]: r for r in self.panel}
        r = hr + 1
        for d in self.dates:
            drv = self.drv.get(d, {})
            _label(ws.cell(row=r, column=1), d, bold=False)
            _val(ws.cell(row=r, column=2), self.h.num(self.fam, "INDEX", SPREAD_METRIC, d), fmt="0.0", color=C_HARD)
            for j, c in enumerate(DRIVER_COLS, start=3):
                _val(ws.cell(row=r, column=j), drv.get(c) if drv.get(c) is not None else "",
                     fmt="0.00", color=C_HARD)
            _val(ws.cell(row=r, column=6), drv.get("regime") if drv.get("regime") is not None else "", fmt="0")
            p = by_date.get(d)
            for j, key in enumerate(["d_spread", "ust10y", "dxy", "vix"], start=7):
                _val(ws.cell(row=r, column=j), p[key] if p else "", fmt="+0.00;-0.00")
            c = ws.cell(row=r, column=11, value="yes" if d in used else "no")
            c.border = BORDER
            _font(c, color=C_FORMULA if d in used else C_NOTE, italic=(d not in used))
            r += 1
        ws.freeze_panes = "B6"

        r += 1
        for line in [
            f"{len(self.dates)} observation(s) of {self.fam} in the store; "
            f"{len(self.drv)} row(s) in {DRIVERS_FILE}; {len(self.panel)} usable monthly change(s).",
            "",
            "The regime column is deliberately NOT derived from VIX. Splitting the sample on a",
            "variable that is already a regressor manufactures a difference between the two betas",
            "even when none exists — tested on data with one true beta of +0.30 throughout, a VIX",
            "split reported +0.47 against -0.14. Set it from something outside the model: real",
            "yields versus breakevens, the equity direction, or your own read.",
        ]:
            ws.cell(row=r, column=1, value=line)
            _font(ws.cell(row=r, column=1), italic=True, color=C_NOTE)
            r += 1

    # ---------------- regression (fitted in Python) ----------------
    def regression(self):
        ws = self.wb.create_sheet("Regression")
        ws.sheet_view.showGridLines = False
        ws["A1"] = f"Spread regression — {self.fam}"
        _font(ws["A1"], size=14, bold=True, color=C_HDR_BG)
        ws["A2"] = ("Monthly CHANGES, not levels: spreads are close to a random walk, so a levels "
                    "regression reports a high R-squared that means nothing.")
        _font(ws["A2"], italic=True, color=C_NOTE)
        ws["A3"] = (f"Fitted by embig_history.py on {datetime.now():%Y-%m-%d %H:%M}. These are values, "
                    "not formulas — rerun the script to re-estimate.")
        _font(ws["A3"], italic=True, color=C_NOTE)
        ws.column_dimensions["A"].width = 38
        for c in "BCDEFG":
            ws.column_dimensions[c].width = 15

        res, fit = self.reg, self.reg.get("fit")
        r = 5
        _label(ws.cell(row=r, column=1), "Sample", fill=C_BAND)
        for c in range(2, 4):
            ws.cell(row=r, column=c).fill = PatternFill("solid", fgColor=C_BAND)
            ws.cell(row=r, column=c).border = BORDER
        r += 1
        stats = [("Specification", SPEC, "General"),
                 ("Usable monthly observations", res["panel_n"], "0")]
        if fit:
            stats += [("R-squared", fit["r2"], "0.000"),
                      ("Adjusted R-squared", fit["adj_r2"], "0.000"),
                      ("Durbin-Watson (2.0 = no autocorrelation)", fit["dw"], "0.00"),
                      ("Std error of the regression (bp)", fit["se_reg"], "0.0"),
                      ("Monthly sigma of spread changes (bp)", fit["sigma_y"], "0.0"),
                      ("12-month sigma (bp)", fit["sigma_y"] * math.sqrt(12), "0.0")]
        for lbl, v, fmt in stats:
            _label(ws.cell(row=r, column=1), lbl, bold=False)
            _val(ws.cell(row=r, column=2), v, fmt=fmt, color=C_HARD, bold=True)
            r += 1

        if not fit:
            r += 1
            for line in [
                f"NOT FITTED — {res['panel_n']} usable monthly change(s); at least 5 are needed to",
                "estimate anything and 24 before a coefficient should be believed.",
                "",
                f"To fit: put monthly {self.fam} spread history in the store (drop the JPM files in",
                f"the folder), put matching UST 10y / DXY / VIX in {DRIVERS_FILE}, and rerun.",
                "Until then the Forecast tab runs on the practitioner priors, which is stated there.",
            ]:
                ws.cell(row=r, column=1, value=line)
                _font(ws.cell(row=r, column=1), italic=True, color=C_NOTE)
                r += 1
            return

        r += 1
        _label(ws.cell(row=r, column=1), "Fitted betas — bp of spread per unit of driver", fill=C_BAND)
        for c in range(2, 8):
            ws.cell(row=r, column=c).fill = PatternFill("solid", fgColor=C_BAND)
            ws.cell(row=r, column=c).border = BORDER
        r += 1
        for i, h in enumerate(["Driver", "fitted beta", "std error", "t-stat", "VIF", "prior",
                               "used in forecast", "why"], start=1):
            _hdr(ws.cell(row=r, column=i), h)
        ws.column_dimensions["H"].width = 30
        r += 1
        picked = chosen_betas(res)
        for i, c in enumerate(res["cols"]):
            pk = DRIVER_PRIORS.get(c)
            _label(ws.cell(row=r, column=1), DRIVER_LABELS[c], bold=False)
            _val(ws.cell(row=r, column=2), fit["beta"][i], fmt="0.000", color=C_HARD)
            _val(ws.cell(row=r, column=3), fit["se"][i], fmt="0.000")
            _val(ws.cell(row=r, column=4), fit["t"][i], fmt="+0.00", bold=abs(fit["t"][i]) >= 2)
            v = fit["vif"][i] if fit.get("vif") else None
            _val(ws.cell(row=r, column=5), v if v is not None else "", fmt="0.0",
                 bold=(v is not None and v > 5))
            _val(ws.cell(row=r, column=6), FORECAST_BETAS[pk] if pk else "", fmt="0.00", color=C_NOTE)
            val, why = picked[c]
            cc = ws.cell(row=r, column=7, value=val)
            cc.number_format = "0.000"; cc.border = BORDER
            cc.fill = PatternFill("solid", fgColor=C_INDEX); _font(cc, bold=True)
            ws.cell(row=r, column=8, value=why).border = BORDER
            _font(ws.cell(row=r, column=8), italic=True, color=C_NOTE)
            r += 1
        _label(ws.cell(row=r, column=1), "Rating drift (per notch)", bold=False)
        for col in (2, 3, 4, 5):
            ws.cell(row=r, column=col).border = BORDER
        _val(ws.cell(row=r, column=6), FORECAST_BETAS["rating_notches"], fmt="0.00", color=C_NOTE)
        cc = ws.cell(row=r, column=7, value=FORECAST_BETAS["rating_notches"])
        cc.number_format = "0.000"; cc.border = BORDER
        cc.fill = PatternFill("solid", fgColor=C_INDEX); _font(cc, bold=True)
        ws.cell(row=r, column=8, value="not fitted — no rating-action series").border = BORDER
        _font(ws.cell(row=r, column=8), italic=True, color=C_NOTE)
        r += 1
        _label(ws.cell(row=r, column=1), "Intercept (bp/month)", bold=False)
        _val(ws.cell(row=r, column=2), fit["intercept"], fmt="+0.00", color=C_HARD)
        _val(ws.cell(row=r, column=3), fit["intercept_se"], fmt="0.00")
        r += 2

        for line in [
            f"A fitted beta is used only when n >= {MIN_OBS_FOR_REGRESSION} AND |t| >= 2. "
            "Otherwise the prior stands; the 'why' column says which rule fired.",
            "VIF above 5 means level and slope are splitting the same variation — the joint fit "
            "can be fine while neither t-stat looks significant. Read them together, not apart.",
            "Durbin-Watson far from 2 means the residuals are autocorrelated and the standard "
            "errors above are too small; treat the t-stats as optimistic.",
        ]:
            ws.cell(row=r, column=1, value=line)
            _font(ws.cell(row=r, column=1), italic=True, color=C_NOTE)
            r += 1
        r += 1

        _label(ws.cell(row=r, column=1), "Conditional UST beta", fill=C_BAND)
        for c in range(2, 6):
            ws.cell(row=r, column=c).fill = PatternFill("solid", fgColor=C_BAND)
            ws.cell(row=r, column=c).border = BORDER
        r += 1
        ws.cell(row=r, column=1, value=("One unconditional beta cannot represent both regimes: rates "
                                        "rising on growth tightens spreads, rates rising on an "
                                        "inflation or policy shock widens them."))
        _font(ws.cell(row=r, column=1), italic=True, color=C_NOTE)
        r += 1
        for i, h in enumerate(["Regime", "n", "beta (UST)", "std error", "t-stat"], start=1):
            _hdr(ws.cell(row=r, column=i), h)
        r += 1
        cond = res.get("conditional", {})
        if not cond:
            ws.cell(row=r, column=1, value="No regime column supplied in drivers.csv — block skipped "
                                           "rather than guessed.")
            _font(ws.cell(row=r, column=1), italic=True, color=C_NOTE)
            r += 1
        for key, lbl in (("growth", "Growth-driven (regime = 0)"), ("shock", "Shock-driven (regime = 1)")):
            if key not in cond:
                continue
            c = cond[key]
            _label(ws.cell(row=r, column=1), lbl, bold=False)
            _val(ws.cell(row=r, column=2), c["n"], fmt="0")
            _val(ws.cell(row=r, column=3), c["beta"] if c["beta"] is not None else "", fmt="0.000",
                 color=C_HARD, bold=True)
            _val(ws.cell(row=r, column=4), c["se"] if c["se"] is not None else "", fmt="0.000")
            _val(ws.cell(row=r, column=5), c["t"] if c["t"] is not None else "", fmt="+0.00")
            r += 1
        b_g = cond.get("growth", {}).get("beta")
        b_s = cond.get("shock", {}).get("beta")
        r += 1
        if b_g is not None and b_s is not None and b_g * b_s < 0:
            ws.cell(row=r, column=1, value=("THE TWO BETAS DIFFER IN SIGN. The single coefficient on the "
                                            "Forecast tab is not describing anything — build the forecast "
                                            "regime by regime."))
            _font(ws.cell(row=r, column=1), bold=True, color="A83A32")
        else:
            ws.cell(row=r, column=1, value=("Split on the exogenous regime flag, never on VIX — VIX is a "
                                            "regressor, and conditioning on it manufactures a difference "
                                            "between the two betas even when none exists."))
            _font(ws.cell(row=r, column=1), italic=True, color=C_NOTE)

    # ---------------- scenarios ----------------
    def scenarios(self):
        ws = self.wb.create_sheet("Scenarios")
        ws.sheet_view.showGridLines = False
        ws["A1"] = f"Total return grid — {self.fam}, {self.d}"
        _font(ws["A1"], size=14, bold=True, color=C_HDR_BG)
        ws["A2"] = ("12m total return = carry less the duration-weighted rate and spread effects. "
                    "Green clears the cash hurdle, red does not.")
        _font(ws["A2"], italic=True, color=C_NOTE)
        ws.column_dimensions["A"].width = 26
        for c in "BCDEFG":
            ws.column_dimensions[c].width = 13

        y = self.h.num(self.fam, "INDEX", YIELD_METRIC, self.d) or 0.0
        ds = self.h.num(self.fam, "INDEX", "Spread Duration", self.d) or 0.0
        di = self.h.num(self.fam, "INDEX", "IR Duration to Worst", self.d) or 0.0
        s0 = self.h.num(self.fam, "INDEX", SPREAD_METRIC, self.d) or 0.0

        r = 4
        for lbl, v, fmt in [("Carry (index yield, %)", y, "0.00"), ("Spread duration", ds, "0.00"),
                            ("IR duration", di, "0.00"), ("Spot spread (bp)", s0, "0"),
                            ("Cash hurdle (%)", 4.50, "0.00")]:
            _label(ws.cell(row=r, column=1), lbl, bold=False)
            c = ws.cell(row=r, column=2, value=v)
            c.number_format = fmt; c.border = BORDER
            if "hurdle" in lbl:
                c.fill = PatternFill("solid", fgColor=C_INPUT)
            _font(c, bold=True, color=C_HARD)
            r += 1
        y_ref, ds_ref, di_ref, s0_ref, cash_ref = "B4", "B5", "B6", "B7", "B8"
        r += 1

        _label(ws.cell(row=r, column=1), "12m total return (%)", fill=C_BAND)
        usts = [0, 25, 50, 75]
        for j, u in enumerate(usts):
            _hdr(ws.cell(row=r, column=2 + j), f"10y {u:+d}bp")
        grid_hdr = r
        r += 1
        spreads = [150, 175, 200, 208, 225, 250, 275]
        for sp in spreads:
            lbl = f"{sp}bp"
            if sp == 175:
                lbl += "   <- your view"
            elif sp == 250:
                lbl += "   <- official"
            _label(ws.cell(row=r, column=1), lbl, bold=(sp in (175, 250)))
            for j, u in enumerate(usts):
                c = ws.cell(row=r, column=2 + j,
                            value=f"={y_ref}-{di_ref}*{u}/100-{ds_ref}*(({sp}-{s0_ref})/100)")
                c.number_format = "0.00"; c.border = BORDER
                _font(c, bold=(sp in (175, 250)))
            r += 1
        r += 1
        _label(ws.cell(row=r, column=1), "Break-even spread vs cash (bp)", bold=False)
        for j, u in enumerate(usts):
            c = ws.cell(row=r, column=2 + j,
                        value=f"={s0_ref}+({y_ref}-{di_ref}*{u}/100-{cash_ref})*100/{ds_ref}")
            c.number_format = "0"; c.border = BORDER; _font(c, bold=True, color=C_HARD)
        r += 2
        ws.cell(row=r, column=1, value=("Read the break-even row first: it is the spread at which this "
                                        "recommendation stops beating cash, which is the number that "
                                        "actually decides the call."))
        _font(ws.cell(row=r, column=1), italic=True, color=C_NOTE)

    # ---------------- forecast ----------------
    def forecast(self):
        ws = self.wb.create_sheet("Forecast")
        ws.sheet_view.showGridLines = False
        ws["A1"] = f"{self.fam} — spread forecast"
        _font(ws["A1"], size=14, bold=True, color=C_HDR_BG)
        ws["A2"] = ("Yellow cells are inputs. Betas are practitioner priors, not regression "
                    "estimates — see the note at the bottom.")
        _font(ws["A2"], italic=True, color=C_NOTE)
        ws.column_dimensions["A"].width = 40
        for c in "BCDEF":
            ws.column_dimensions[c].width = 15

        cur = stripped(self.h, self.fam, "INDEX", self.d, self.theta) if self.theta \
            else self.h.num(self.fam, "INDEX", SPREAD_METRIC, self.d)
        cur = cur or 0.0

        r = 4
        _label(ws.cell(row=r, column=1), "Current state", fill=C_BAND)
        ws.cell(row=r, column=2).fill = PatternFill("solid", fgColor=C_BAND)
        ws.cell(row=r, column=2).border = BORDER
        r += 1
        for lbl, val, fmt in [("Spread today (bp)", cur, "0.0"),
                              ("Your 12m view (bp) — edit", 175.0, "0"),
                              ("Index yield (%)", self.h.num(self.fam, "INDEX", YIELD_METRIC, self.d), "0.00"),
                              ("Spread duration", self.h.num(self.fam, "INDEX", "Spread Duration", self.d), "0.00"),
                              ("IR duration", self.h.num(self.fam, "INDEX", "IR Duration to Worst", self.d), "0.00"),
                              ("YTD total return (%)", self.h.num(self.fam, "INDEX", "YTD Change (%)", self.d), "+0.00")]:
            _label(ws.cell(row=r, column=1), lbl, bold=False)
            if lbl.startswith("Your 12m view"):
                c = ws.cell(row=r, column=2, value=val)
                c.number_format = fmt; c.border = BORDER
                c.fill = PatternFill("solid", fgColor=C_INPUT); _font(c, bold=True)
                view_ref = f"B{r}"
            else:
                _val(ws.cell(row=r, column=2), val if val is not None else "", fmt=fmt, color=C_HARD)
            r += 1

        _label(ws.cell(row=r, column=1), "12m total return on that view (%)", bold=False)
        c = ws.cell(row=r, column=2, value=f"=B7-B8*(({view_ref}-B5)/100)")
        c.number_format = "0.00"; c.border = BORDER; _font(c, bold=True, color=C_HARD)
        ws.cell(row=r, column=3, value="rates unchanged; see the Scenarios tab for the rate grid")
        _font(ws.cell(row=r, column=3), italic=True, color=C_NOTE)
        r += 2

        _label(ws.cell(row=r, column=1), "Driver views (edit the yellow cells)", fill=C_BAND)
        for c in range(2, 4):
            ws.cell(row=r, column=c).fill = PatternFill("solid", fgColor=C_BAND)
            ws.cell(row=r, column=c).border = BORDER
        r += 1
        _hdr(ws.cell(row=r, column=1), "Driver")
        _hdr(ws.cell(row=r, column=2), "6m view")
        _hdr(ws.cell(row=r, column=3), "12m view")
        _hdr(ws.cell(row=r, column=4), "beta (bp of spread)")
        _hdr(ws.cell(row=r, column=5), "source")
        r += 1
        drv_start = r
        picked = chosen_betas(self.reg)
        blocks = [(DRIVER_LABELS[c] + " change", c) for c in (self.reg.get("cols") or DRIVER_COLS)]
        blocks.append(("Rating drift (notches)", None))
        for lbl, key in blocks:
            _label(ws.cell(row=r, column=1), lbl, bold=False)
            for col in (2, 3):
                c = ws.cell(row=r, column=col, value=0.0)
                c.number_format = "0.0"; c.border = BORDER
                c.fill = PatternFill("solid", fgColor=C_INPUT)
                _font(c, bold=True)
            if key is None:
                val, why = FORECAST_BETAS["rating_notches"], "prior — not fitted"
            else:
                val, why = picked.get(key, (0.0, "prior"))
            c = ws.cell(row=r, column=4, value=val)
            c.number_format = "0.000"; c.border = BORDER; _font(c, color=C_NOTE)
            c = ws.cell(row=r, column=5, value=why)
            c.border = BORDER; _font(c, italic=True, color=C_NOTE, size=9)
            r += 1
        drv_end = r - 1

        r += 1
        _label(ws.cell(row=r, column=1), "Vol multiplier (1.0 = historical)", bold=False)
        vm = ws.cell(row=r, column=2, value=1.0)
        vm.number_format = "0.00"; vm.border = BORDER
        vm.fill = PatternFill("solid", fgColor=C_INPUT); _font(vm, bold=True)
        vm_ref = f"B{r}"
        r += 2

        # sigma from this family's own spread history
        ser = [v for _, v in self.h.series(self.fam, "INDEX", SPREAD_METRIC)]
        if len(ser) >= 3:
            diffs = [ser[i + 1] - ser[i] for i in range(len(ser) - 1)]
            mu = sum(diffs) / len(diffs)
            sigma = math.sqrt(sum((x - mu) ** 2 for x in diffs) / len(diffs))
            sig_note = f"from {len(ser)} observations of this family"
        else:
            sigma = 45.0
            sig_note = ("DEFAULT 45bp — this family has fewer than 3 observations. "
                        "It self-calibrates as you accumulate downloads.")
        _label(ws.cell(row=r, column=1), "Monthly sigma of spread changes (bp)", bold=False)
        fit = self.reg.get("fit")
        if fit:
            sigma, sig_note = fit["sigma_y"], f"fitted from {fit['n']} monthly changes"
        c = ws.cell(row=r, column=2, value=sigma)
        c.number_format = "0.0"; c.border = BORDER; _font(c, bold=True, color=C_HARD)
        ws.cell(row=r, column=3, value=sig_note)
        sig_m = f"B{r}"
        r += 1
        _label(ws.cell(row=r, column=1), "12-month sigma (bp)", bold=False)
        c = ws.cell(row=r, column=2, value=f"={sig_m}*SQRT(12)")
        c.number_format = "0.0"; c.border = BORDER; _font(c, bold=True)
        ws.cell(row=r, column=3, value="monthly sigma scaled by root-12 — the bands are a 12m horizon")
        _font(ws.cell(row=r, column=3), italic=True, color=C_NOTE)
        sig_ref = f"B{r}"
        r += 2

        _label(ws.cell(row=r, column=1), "Forecast", fill=C_BAND)
        for c in range(2, 5):
            ws.cell(row=r, column=c).fill = PatternFill("solid", fgColor=C_BAND)
            ws.cell(row=r, column=c).border = BORDER
        r += 1
        _hdr(ws.cell(row=r, column=1), "Horizon")
        _hdr(ws.cell(row=r, column=2), "Spread (bp)")
        _hdr(ws.cell(row=r, column=3), "Change (bp)")
        _hdr(ws.cell(row=r, column=4), "Implied 12m TR (%)")
        r += 1

        # 3m = quarter of the 6m view; 6m and 12m as given
        cur_ref = "B5"
        for lbl, scale, col in [("3 months", 0.5, 2), ("6 months", 1.0, 2), ("12 months", 1.0, 3)]:
            terms = "+".join(f"{get_column_letter(col)}{i}*D{i}" for i in range(drv_start, drv_end + 1))
            delta = f"({terms})*{scale}"
            _label(ws.cell(row=r, column=1), lbl, bold=False)
            c = ws.cell(row=r, column=2, value=f"={cur_ref}+{delta}")
            c.number_format = "0.0"; c.border = BORDER; _font(c)
            c = ws.cell(row=r, column=3, value=f"={delta}")
            c.number_format = "+0.0;-0.0"; c.border = BORDER; _font(c)
            c = ws.cell(row=r, column=4, value=f"=B7-B9*({delta})/100")
            c.number_format = "0.00"; c.border = BORDER; _font(c)
            r += 1

        r += 1
        _label(ws.cell(row=r, column=1), "12m percentile bands (bp)", fill=C_BAND)
        for c in range(2, 7):
            ws.cell(row=r, column=c).fill = PatternFill("solid", fgColor=C_BAND)
            ws.cell(row=r, column=c).border = BORDER
        r += 1
        for i, p in enumerate(FORECAST_PERCENTILES):
            _hdr(ws.cell(row=r, column=2 + i), f"p{p}")
        _label(ws.cell(row=r, column=1), "Gaussian around the 12m central case", bold=False)
        r += 1
        terms12 = "+".join(f"C{i}*D{i}" for i in range(drv_start, drv_end + 1))
        for i, p in enumerate(FORECAST_PERCENTILES):
            z = {5: -1.645, 25: -0.674, 50: 0.0, 75: 0.674, 95: 1.645}[p]
            c = ws.cell(row=r, column=2 + i, value=f"={cur_ref}+({terms12})+{z}*{sig_ref}*{vm_ref}")
            c.number_format = "0"; c.border = BORDER; _font(c, bold=(p == 50))
        r += 3

        for line in [
            "Betas are practitioner priors, not in-sample estimates. Known weakness: beta(UST) is",
            "unconditional, so it cannot distinguish rates rising on growth (spreads tighten) from",
            "rates rising on an inflation shock (spreads widen). Judge it accordingly.",
            "",
            "Sigma is estimated from THIS family's spread history only. Observations from another",
            "index family are never mixed in, so the number is honest but small until the store fills.",
        ]:
            ws.cell(row=r, column=1, value=line)
            _font(ws.cell(row=r, column=1), italic=True, color=C_NOTE)
            r += 1

    # ---------------- methodology ----------------
    def methodology(self):
        ws = self.wb.create_sheet("Methodology")
        ws.sheet_view.showGridLines = False
        ws.column_dimensions["A"].width = 118
        ws["A1"] = "Methodology and known limitations"
        _font(ws["A1"], size=14, bold=True, color=C_HDR_BG)
        fit = self.reg.get("fit")
        anchor = SPREAD_ANCHORS.get(self.fam)
        blocks = [
            ("Index family", [
                f"This workbook is {self.fam} only. EMBI Global and EMBI Global Diversified hold the",
                "same bonds under different weights - Diversified caps the larger-debt issuers - so the",
                "two are stored and reported separately and are never joined into one series. Joining",
                "them books index-change artefacts as market moves.",
            ]),
            ("Spread metric", [
                f"{SPREAD_METRIC}: duration-weighted spread to worst over the Treasury curve.",
                "No JPM file we receive carries a stripped spread. STW is measured off the par curve,",
                "Z-spread off the zero curve, and the published stripped spread sits between them.",
                "spread* = STW + theta x (Z - STW), with theta solved ONCE against a published",
                "reference on a known date, then applied unchanged to every entity and every date.",
                (f"Anchor for this family: {anchor[1]}bp on {anchor[0]}, theta {self.theta:.4f}."
                 if anchor and self.theta else
                 "No anchor set for this family - raw STW is shown and the proxy column is blank."),
                "Where Z < STW the measures have inverted (defaulted paper whose yield-to-worst is an",
                "artefact) and those rows stay on raw STW rather than being dragged the wrong way.",
            ]),
            ("Which Treasury drives the spread", [
                "Duration is a sensitivity, not a point on the curve. The index shows ~6.2 duration but",
                "10.2y average life; duration is shorter only because coupons pay along the way.",
                "Backing the Treasury anchor out of each country's own numbers (implied UST = yield -",
                "STW) and regressing across 66 countries: average life fits at R2 0.88 against 0.84 for",
                "duration, and the index anchor of 4.916% sits at an 11.2y point. The 5y point is 30bp",
                "below where the index is actually struck, so the 10y is the level factor.",
                "The default specification adds the 2s10s slope, because a twist is a different event",
                "from a parallel shift. Whether slope earns its degree of freedom is decided by",
                "ADJUSTED R-squared (--spec compare), never by plain R-squared, which can only rise.",
                "Hedging is a separate question: a 6.2-duration book against an ~8-duration 10y note is",
                "a 0.78 ratio, and a 5s/10s blend is a legitimate way to build it.",
            ]),
            ("Regression", [
                "Monthly CHANGES, not levels. Spreads are close to a random walk, so a levels",
                "regression reports a flattering R-squared that means nothing.",
                f"A fitted beta is used only when n >= {MIN_OBS_FOR_REGRESSION} and |t| >= 2; otherwise",
                "the practitioner prior stands and the Regression tab says which rule fired.",
                "Drivers with no prior (the slope factor) contribute zero unless they are significant.",
                (f"Current fit: n={fit['n']}, R2={fit['r2']:.3f}, adjusted {fit['adj_r2']:.3f}, "
                 f"DW {fit['dw']:.2f}, residual se {fit['se_reg']:.1f}bp."
                 if fit else
                 f"NOT FITTED - {self.reg['panel_n']} usable monthly change(s). Priors in use."),
            ]),
            ("Known weaknesses - state these before anyone else does", [
                "1. beta(UST) is unconditional. It cannot distinguish rates rising on growth (spreads",
                "   tighten) from rates rising on an inflation or policy shock (spreads widen). The",
                "   conditional block splits the sample on an EXOGENOUS regime flag you set by hand.",
                "   It is deliberately not derived from VIX: VIX is already a regressor, and",
                "   conditioning on it manufactures a regime difference even when none exists - tested",
                "   on data with one true beta of +0.30 throughout, a VIX split reported +0.47 vs -0.14.",
                "2. Rating drift is never fitted - there is no rating-action series in these files.",
                "3. Percentile bands are Gaussian around the central case, with 12m sigma taken as the",
                "   monthly sigma scaled by root-12. Fat tails are not modelled; the vol multiplier on",
                "   the Forecast tab is the crude lever for that.",
                "4. Theta is calibrated on ONE date and holds only while the par/zero wedge is stable.",
                "   Re-anchor monthly; if it drifts outside roughly 0.5-0.7, stop interpolating.",
                "5. Everything here is duration-weighted because that is what JPM publishes. A",
                "   market-value-weighted aggregate is a different number.",
            ]),
        ]
        r = 3
        for title, lines in blocks:
            _label(ws.cell(row=r, column=1), title, fill=C_BAND)
            r += 1
            for ln in lines:
                ws.cell(row=r, column=1, value=ln)
                _font(ws.cell(row=r, column=1), color=C_NOTE)
                r += 1
            r += 1

    # ---------------- raw ----------------
    def data_raw(self):
        ws = self.wb.create_sheet("Store_Dump")
        ws.sheet_view.showGridLines = False
        ws["A1"] = "Every row held for this family (the store itself lives in embig_history.csv)"
        _font(ws["A1"], size=12, bold=True, color=C_HDR_BG)
        hr = 3
        for i, h in enumerate(["date", "section", "entity", "metric", "value"], start=1):
            _hdr(ws.cell(row=hr, column=i), h)
            ws.column_dimensions[get_column_letter(i)].width = [12, 12, 28, 30, 16][i - 1]
        r = hr + 1
        for (d, fam, sec, ent, met), v in sorted(self.h.data.items()):
            if fam != self.fam:
                continue
            for i, val in enumerate([d, sec, ent, met, v], start=1):
                c = ws.cell(row=r, column=i, value=_num(val) if i == 5 and _num(val) is not None else val)
                c.border = BORDER; _font(c)
            r += 1
        ws.freeze_panes = "A4"

    def build(self) -> Workbook:
        self.cover()
        self.country_spreads()
        self.performance()
        self.spread_history()
        self.scenarios()
        self.drivers()
        self.regression()
        self.forecast()
        self.methodology()
        self.data_raw()
        return self.wb


# ===========================================================================
# 7. DRIVER
# ===========================================================================

def main(argv: Optional[List[str]] = None) -> int:
    global SPEC, DRIVER_COLS
    ap = argparse.ArgumentParser(description="Ingest JPM EMBI downloads and build the dashboard.")
    ap.add_argument("--dir", default=".", help="folder to scan (default: current)")
    ap.add_argument("--output", "-o", default=DEFAULT_OUTPUT)
    ap.add_argument("--family", default=None, help="which index family to build (default: the one with the latest data)")
    ap.add_argument("--no-build", action="store_true", help="ingest only, do not write the workbook")
    ap.add_argument("--list", action="store_true", help="show what the store holds and exit")
    ap.add_argument("--spec", choices=list(RATE_SPECS) + ["compare"], default=SPEC,
                    help="rate specification: level (10y only), level_slope (10y + 2s10s, "
                         "default), front (5y), or compare to fit all three")
    a = ap.parse_args(argv)

    folder = Path(a.dir).expanduser().resolve()
    store_path = folder / HISTORY_FILE
    hist = History()
    hist.load(store_path)

    if a.list:
        if not hist.data:
            print("Store is empty.")
            return 0
        print(f"{store_path}  —  {len(hist.data):,} observations")
        for fam in hist.families():
            ds = hist.dates(fam)
            print(f"\n  {fam}")
            print(f"    {len(ds)} dates: {ds[0]} .. {ds[-1]}")
            print(f"    countries {len(hist.entities(fam,'country'))}, "
                  f"regions {len(hist.entities(fam,'region'))}, "
                  f"rating buckets {len(hist.entities(fam,'subcredit'))}")
            sp = hist.num(fam, "INDEX", SPREAD_METRIC, ds[-1])
            print(f"    latest index {SPREAD_METRIC}: {sp}")
        return 0

    files = sorted(p for p in folder.glob("*.csv") if p.name != HISTORY_FILE)
    archive = folder / ARCHIVE_DIR
    if archive.exists():
        files += sorted(archive.rglob("*.csv"))

    ingested = 0
    print(f"Scanning {folder}")
    for p in files:
        if not is_jpm_snapshot(p):
            continue
        snap = parse_snapshot(p)
        if not snap:
            print(f"  [skip]     {p.name} — no usable rows")
            continue
        added, updated = hist.ingest(snap)
        ingested += 1
        d = snap["date"].strftime("%Y-%m-%d")
        print(f"  [{snap['family']:<26}] {d}  {p.name}  (+{added} new, {updated} revised)")
        # archive under the family so a rebuild from raw is always possible
        dest_dir = archive / snap["family"].replace(" ", "_")
        dest_dir.mkdir(parents=True, exist_ok=True)
        dest = dest_dir / f"snapshot_{d}.csv"
        if p.resolve() != dest.resolve() and not dest.exists():
            dest.write_bytes(p.read_bytes())
            print(f"                                 archived -> {dest.relative_to(folder)}")

    if not hist.data:
        print("ERROR: no JPM snapshot files found and the store is empty.", file=sys.stderr)
        return 1

    hist.save(store_path)
    print(f"\nStore: {len(hist.data):,} observations across {len(hist.families())} "
          f"famil{'y' if len(hist.families())==1 else 'ies'} -> {store_path.name}")
    for fam in hist.families():
        ds = hist.dates(fam)
        print(f"   {fam:<28} {len(ds):>3} date(s)  {ds[0]} .. {ds[-1]}")

    if a.no_build:
        return 0

    family = a.family
    if family and family not in hist.families():
        print(f"ERROR: no data for family {family!r}. Have: {hist.families()}", file=sys.stderr)
        return 1
    if not family:
        family = max(hist.families(), key=lambda f: hist.dates(f)[-1])

    if len(hist.families()) > 1:
        print(f"\nNOTE: the store holds more than one index family. Building {family!r} only.")
        print("      Families are never merged — rerun with --family to build the other.")
    if SPREAD_ANCHORS.get(family) is None:
        print(f"NOTE: no stripped-spread anchor set for {family!r}; showing raw {SPREAD_METRIC}.")
        print("      Add one to SPREAD_ANCHORS once you have a published reference print.")

    if a.spec != "compare":
        SPEC = a.spec
    DRIVER_COLS = RATE_SPECS[SPEC]

    drv_path = folder / DRIVERS_FILE
    drivers = load_drivers(drv_path)
    if not drv_path.exists():
        write_drivers_template(drv_path)
        print(f"\nCreated {DRIVERS_FILE} — fill in UST 10y / DXY / VIX per month to fit the regression.")
    else:
        print(f"\nDrivers: {len(drivers)} row(s) from {DRIVERS_FILE}")

    dash = Dashboard(hist, family, drivers)
    if dash.reg.get("fit"):
        f = dash.reg["fit"]
        print(f"Regression [{SPEC}]: n={f['n']}, R2={f['r2']:.3f}, adjR2={f['adj_r2']:.3f}, "
              f"DW={f['dw']:.2f}, sigma={f['sigma_y']:.1f}bp/month")
        for i, c in enumerate(dash.reg["cols"]):
            vif = f["vif"][i] if f.get("vif") else float("nan")
            print(f"   {DRIVER_LABELS[c]:<24} beta {f['beta'][i]:+7.3f}  se {f['se'][i]:.3f}  "
                  f"t {f['t'][i]:+6.2f}  VIF {vif:4.1f}")
        for k, v in chosen_betas(dash.reg).items():
            print(f"   -> forecast uses {DRIVER_LABELS[k]:<20} {v[0]:+7.3f}   ({v[1]})")
        for k, cd in dash.reg.get("conditional", {}).items():
            if cd.get("beta") is not None:
                print(f"   conditional UST beta, {k:<7} n={cd['n']:<3} {cd['beta']:+7.3f}  t {cd['t']:+6.2f}")
    else:
        print(f"Regression: not fitted — {dash.reg['panel_n']} usable monthly change(s); need 5+ "
              f"({MIN_OBS_FOR_REGRESSION}+ to be believed). Forecast runs on priors.")

    if a.spec == "compare":
        cmp = compare_specs(dash.panel)
        print("\n  Specification comparison — adjusted R2 is the one to read:")
        best = None
        for name in RATE_SPECS:
            f = cmp.get(name)
            if not f:
                print(f"    {name:<12} not enough data")
                continue
            mx = max(f["vif"]) if f.get("vif") else float("nan")
            print(f"    {name:<12} n={f['n']:<4} R2 {f['r2']:.3f}  adjR2 {f['adj_r2']:.3f}  "
                  f"resid se {f['se_reg']:.1f}bp  maxVIF {mx:.1f}")
            if best is None or f["adj_r2"] > cmp[best]["adj_r2"]:
                best = name
        if best:
            print(f"    -> '{best}' wins on adjusted R2. Rerun with --spec {best}.")
        print("    A max VIF above 5 means level and slope are splitting the same variation;")
        print("    the joint fit can still be fine while neither t-stat looks significant.")
        print("    The 0.30 prior was calibrated on 10y moves — it does not carry to 'front'.")

    wb = dash.build()
    out = Path(a.output).expanduser()
    if not out.is_absolute():
        out = folder / out
    wb.save(out)
    print(f"\nWrote {out}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
