# -*- coding: utf-8 -*-
"""
best_bonds_by_country.py
========================
Finds the best bonds on each country curve, judged on the KPI that matters:
EXPECTED 12-MONTH TOTAL RETURN vs. THE MEDIAN OF SIMILARLY RATED PEERS.

Run:   F5 in Spyder (edit CONFIG), or
       python best_bonds_by_country.py
       python best_bonds_by_country.py --country Peru --country Colombia
       python best_bonds_by_country.py --data "data/current/Offshore_EMBL.xlsx"
       python best_bonds_by_country.py --show-columns      (check column detection)

This file also holds the ENGINE. best_bonds_by_issuer.py imports it, so both
scripts always score bonds the same way.

-----------------------------------------------------------------------------
MAINTAINER'S GUIDE (where to change things)
-----------------------------------------------------------------------------
  Countries to run ............ CONFIG["COUNTRIES"]
  Input file / folder ......... CONFIG["DATA_PATH"]
  A column is not detected .... CONFIG["COLUMN_MAP"]  (your header -> field)
  Peer definition ............. PEER_NOTCH_BAND / PEER_MATURITY_BAND / PEER_MIN_COUNT
  Scenario size / betas ....... SCENARIO_SHIFT_BP / SCENARIO_BETA
  How much cheapness closes ... REVERSION_SHARE   (0 = ignore curve cheapness)
  Default / recovery table .... ANNUAL_PD_BY_NOTCH / RECOVERY_RATE
  Score weights ............... SCORE_WEIGHTS
  Curve buckets ............... BUCKETS
-----------------------------------------------------------------------------
"""

import argparse
import math
import os
import re
import sys
from datetime import date, datetime

import numpy as np
import pandas as pd

# =============================================================================
# CONFIG
# =============================================================================
CONFIG = {
    # --- what to run -------------------------------------------------------
    "COUNTRIES": ["Argentina", "Chile", "Mexico", "Brazil"],
    "UNIVERSE": "sovereign",          # "sovereign" | "all" (sovereign + quasi + corporates, one curve per issuer)
    "CURRENCY": "USD",                # "all" to keep every currency (peers always matched on currency)

    # --- input -------------------------------------------------------------
    # A file (.xlsx/.xls/.csv/.txt) or a folder (newest matching file is used).
    "DATA_PATH": "data/current",
    "FILE_PATTERN": r"(?i)offshore.*\.(xlsx|xls)$|\.(csv|txt)$",
    "SHEET": None,                    # None = auto-pick the sheet that looks like a bond list
    "AS_OF": None,                    # "2026-09-17" or None = today
    # Force a column mapping if auto-detection misses one: {"Your Header": "field"}
    # fields: isin issuer country currency coupon maturity price yield sp moody fitch
    #         duration view amount issuer_type seniority call_date
    "COLUMN_MAP": {},

    # --- output ------------------------------------------------------------
    "OUTPUT_DIR": "outputs",
    "OUTPUT_NAME": None,              # None = Best_Bonds_<countries>_<date>.xlsx

    # --- return model ------------------------------------------------------
    "HORIZON_YEARS": 1.0,
    "MIN_YEARS_TO_MATURITY": 1.25,    # bonds maturing inside the horizon aren't comparable
    "REVERSION_SHARE": 0.5,           # share of curve cheapness/richness assumed to close in 12m
    "RECOVERY_RATE": 0.40,
    "USE_EXPECTED_LOSS": False,       # KPI is market total return; True deducts rating-based expected default loss

    # --- curve fit ---------------------------------------------------------
    "CURVE_MIN_BONDS": 4,             # fit a+b*ln(1+t)+c*t; 3 bonds -> a+b*ln(1+t); <3 -> no curve
    "CURVE_OUTLIER_SIGMA": 2.5,

    # --- peers (the benchmark) --------------------------------------------
    "PEER_NOTCH_BAND": 1,             # composite rating within +/- N notches
    "PEER_MATURITY_BAND": 2.0,        # years to maturity within +/- N years
    "PEER_MATURITY_BAND_WIDE": 3.5,   # used if too few peers at the tight band
    "PEER_NOTCH_BAND_WIDE": 3,        # last resort for thinly populated ratings (e.g. CCC); disclosed per bond
    "PEER_MIN_COUNT": 5,
    "PEER_EXCLUDE_SAME_COUNTRY": True,
    "PEER_EXCLUDE_SELL_VIEW": False,

    # --- scenarios ---------------------------------------------------------
    "SCENARIO_SHIFT_BP": 100,         # EM rally = -shift x beta ; sell-off = +shift x beta
    "SCENARIO_BETA": {"IG": 0.7, "BB": 1.0, "B": 1.4, "CCC": 2.0},

    # --- ranking -----------------------------------------------------------
    "SCORE_WEIGHTS": {"base": 0.50, "rally": 0.20, "selloff": 0.30},
    "MIN_EXCESS_BP": 0,               # a pick must beat the peer median by at least this (base case)
    "TOP_N_PER_CURVE": 3,
    "EXCLUDE_VIEWS": ["Sell"],        # CIO views never picked (still shown on country tabs)
    "SMALL_ISSUE_MM": 500,            # amount outstanding below this -> liquidity penalty
    "SMALL_ISSUE_PENALTY_BP": 25,
    "SUBORDINATED_PENALTY_BP": 50,

    "BUCKETS": [("Short (<3y)", 0, 3), ("Belly (3-7y)", 3, 7),
                ("Intermediate (7-12y)", 7, 12), ("Long (12y+)", 12, 99)],
}

# Approximate long-run 1-year default probabilities by composite notch (%),
# in the spirit of S&P/Moody's multi-decade studies. Edit if you use a house table.
ANNUAL_PD_BY_NOTCH = {
    1: 0.00, 2: 0.01, 3: 0.02, 4: 0.03, 5: 0.05, 6: 0.06, 7: 0.08,
    8: 0.12, 9: 0.17, 10: 0.25, 11: 0.35, 12: 0.55, 13: 0.90,
    14: 1.50, 15: 2.50, 16: 4.00, 17: 7.00, 18: 11.0, 19: 17.0,
    20: 25.0, 21: 35.0, 22: 50.0,
}

COUNTRY_ALIASES = {
    "brasil": "brazil", "mexico": "mexico", "méxico": "mexico",
    "united mexican states": "mexico", "federative republic of brazil": "brazil",
    "republic of chile": "chile", "argentine republic": "argentina",
    "republic of argentina": "argentina",
}

SOVEREIGN_NAME_HINTS = [
    "republic", "united mexican states", "government", "kingdom", "state of",
    "ministry", "treasury", "sultanate", "emirate", "commonwealth", "federation",
    "sovereign",
]

# =============================================================================
# Column detection
# =============================================================================
FIELD_ALIASES = {
    "isin":        ["isin", "isin code", "isin/valor", "security id"],
    "issuer":      ["issuer", "issuer name", "name", "borrower", "company"],
    "country":     ["country", "country of risk", "cntry", "country name", "risk country"],
    "currency":    ["currency", "ccy", "crncy", "cur"],
    "coupon":      ["coupon", "cpn", "coupon (%)", "coupon rate"],
    "maturity":    ["maturity", "maturity date", "mat date", "maturity/call", "final maturity"],
    "price":       ["offer price", "ask price", "price", "px ask", "mid price", "offer px", "price (offer)"],
    "yield":       ["offer yield", "ask yield", "yield", "ytw", "ytm", "yield to worst",
                    "yield to maturity", "yld", "yield (offer)"],
    "sp":          ["s&p", "s&p rating", "sp rating", "rtg_sp", "s&p issue rating"],
    "moody":       ["moody's", "moodys", "moody's rating", "rtg_moody", "moody"],
    "fitch":       ["fitch", "fitch rating", "rtg_fitch"],
    "duration":    ["modified duration", "mod duration", "mod dur", "duration", "dur"],
    "view":        ["cio view", "view", "recommendation", "house view", "cio rating"],
    "amount":      ["amount outstanding", "amt outstanding", "outstanding", "issue size",
                    "amount issued", "size"],
    "issuer_type": ["issuer type", "sector", "type", "asset class", "bond type", "segment"],
    "seniority":   ["seniority", "rank", "payment rank", "subordination", "ranking"],
    "call_date":   ["next call date", "call date", "next call"],
}


def _norm(s):
    return re.sub(r"\s+", " ", str(s).strip().lower())


def _match_headers(headers, column_map):
    """Return {field: original_header}. COLUMN_MAP wins, then exact alias, then contains."""
    out = {}
    normed = {h: _norm(h) for h in headers if h is not None and str(h).strip()}
    for header, field in (column_map or {}).items():
        for h in normed:
            if _norm(h) == _norm(header):
                out[field] = h
    used = set(out.values())
    for field, aliases in FIELD_ALIASES.items():
        if field in out:
            continue
        for alias in aliases:                                   # exact
            hit = next((h for h, n in normed.items() if n == alias and h not in used), None)
            if hit:
                out[field] = hit; used.add(hit); break
    for field, aliases in FIELD_ALIASES.items():
        if field in out or field in ("name",):
            continue
        for alias in aliases:                                   # contains (longer aliases only)
            if len(alias) < 5:
                continue
            hit = next((h for h, n in normed.items() if alias in n and h not in used), None)
            if hit:
                out[field] = hit; used.add(hit); break
    return out


def _find_table(raw, column_map):
    """raw = header-less DataFrame. Find the header row with the most field hits."""
    best = (0, None, {})
    for r in range(min(40, len(raw))):
        headers = [str(x) if pd.notna(x) else None for x in raw.iloc[r].tolist()]
        m = _match_headers(headers, column_map)
        score = len(m) + (3 if "isin" in m else 0) + (2 if "yield" in m or "price" in m else 0)
        if score > best[0]:
            best = (score, r, m)
    return best


def _resolve_path(path, pattern):
    if os.path.isfile(path):
        return path
    if os.path.isdir(path):
        files = [os.path.join(path, f) for f in os.listdir(path)
                 if re.search(pattern, f) and not f.startswith("~$")]
        if not files:
            raise FileNotFoundError(f"No file matching {pattern!r} in {path}")
        return max(files, key=os.path.getmtime)
    raise FileNotFoundError(f"DATA_PATH not found: {path}")


def load_universe(cfg, verbose=True):
    path = _resolve_path(cfg["DATA_PATH"], cfg["FILE_PATTERN"])
    ext = os.path.splitext(path)[1].lower()
    candidates = []
    if ext in (".xlsx", ".xls", ".xlsm"):
        sheets = pd.read_excel(path, sheet_name=None, header=None, dtype=object)
        if cfg.get("SHEET"):
            sheets = {cfg["SHEET"]: sheets[cfg["SHEET"]]}
        for name, raw in sheets.items():
            score, r, m = _find_table(raw, cfg["COLUMN_MAP"])
            if r is not None:
                bonus = 2 if re.search(r"(?i)bond ?list", name) else 0
                candidates.append((score + bonus, name, r, m, raw))
    else:
        sep = "\t" if ext == ".txt" else ","
        raw = pd.read_csv(path, header=None, dtype=object, sep=sep, engine="python")
        score, r, m = _find_table(raw, cfg["COLUMN_MAP"])
        candidates.append((score, os.path.basename(path), r, m, raw))
    if not candidates:
        raise ValueError(f"Could not find a bond table in {path}")
    _, sheet, hdr_row, mapping, raw = max(candidates, key=lambda c: c[0])
    df = raw.iloc[hdr_row + 1:].copy()
    df.columns = [str(x) if pd.notna(x) else f"_col{i}" for i, x in enumerate(raw.iloc[hdr_row])]
    df = df.rename(columns={v: k for k, v in mapping.items()})
    keep = [f for f in FIELD_ALIASES if f in df.columns]
    df = df[keep].copy()
    if verbose:
        print(f"Input : {path}  [sheet: {sheet}, header row {hdr_row + 1}]")
        print("Columns detected: " + ", ".join(f"{k}<-'{v}'" for k, v in mapping.items()))
        missing = [f for f in ("isin", "issuer", "country", "coupon", "maturity", "sp", "moody")
                   if f not in mapping]
        if missing:
            print("  ! not detected: " + ", ".join(missing) + "  (use CONFIG['COLUMN_MAP'])")
    return df, path, mapping


# =============================================================================
# Cleaning
# =============================================================================
SP_SCALE = ["AAA", "AA+", "AA", "AA-", "A+", "A", "A-", "BBB+", "BBB", "BBB-", "BB+", "BB",
            "BB-", "B+", "B", "B-", "CCC+", "CCC", "CCC-", "CC", "C", "D"]
MOODY_SCALE = ["AAA", "AA1", "AA2", "AA3", "A1", "A2", "A3", "BAA1", "BAA2", "BAA3", "BA1", "BA2",
               "BA3", "B1", "B2", "B3", "CAA1", "CAA2", "CAA3", "CA", "C"]


def notch_sp(x):
    if x is None or (isinstance(x, float) and math.isnan(x)):
        return None
    s = str(x).upper().replace("(P)", "").replace("*", "").strip()
    s = re.split(r"[\s/(]", s)[0].rstrip("U").strip()
    if s in ("SD", "RD", "D", "DDD", "DD"):
        return 22
    return SP_SCALE.index(s) + 1 if s in SP_SCALE else None


def notch_moody(x):
    if x is None or (isinstance(x, float) and math.isnan(x)):
        return None
    s = str(x).upper().replace("(P)", "").replace("*", "").strip()
    s = re.split(r"[\s/(]", s)[0]
    s = re.sub(r"(ST|U)$", "", s)
    return MOODY_SCALE.index(s) + 1 if s in MOODY_SCALE else None


def notch_to_label(n):
    if n is None or (isinstance(n, float) and math.isnan(n)):
        return "NR"
    return SP_SCALE[int(round(n)) - 1]


def grade_bucket(n):
    if n <= 10:
        return "IG"
    if n <= 13:
        return "BB"
    if n <= 16:
        return "B"
    return "CCC"


def composite_notch(*notches):
    """Middle of three, lower (worse) of two, the one if one."""
    v = sorted(x for x in notches if x is not None)
    if not v:
        return None
    if len(v) >= 3:
        return v[len(v) // 2]
    return v[-1]


def to_float(x):
    if x is None:
        return None
    if isinstance(x, (int, float, np.floating)):
        return None if (isinstance(x, float) and math.isnan(x)) else float(x)
    s = str(x).replace(",", "").replace("%", "").strip()
    try:
        return float(s)
    except ValueError:
        return None


def to_date(x):
    if x is None or (isinstance(x, float) and math.isnan(x)):
        return None
    if isinstance(x, (datetime, pd.Timestamp)):
        return x.date() if hasattr(x, "date") else x
    if isinstance(x, date):
        return x
    if isinstance(x, (int, float)) and 20000 < x < 80000:     # Excel serial
        return (pd.Timestamp("1899-12-30") + pd.Timedelta(days=float(x))).date()
    s = str(x).strip()
    if s.upper() in ("PERP", "PERPETUAL", ""):
        return None
    try:
        return pd.to_datetime(s, dayfirst=False, errors="raise").date()
    except Exception:
        return None


def norm_country(s):
    n = _norm(s)
    return COUNTRY_ALIASES.get(n, n)


def is_sovereign(row):
    t = _norm(row.get("issuer_type", "") or "")
    if t and t != "nan":
        if any(k in t for k in ("sovereign", "government", "govt")):
            return True
        if any(k in t for k in ("corp", "quasi", "financial", "agency", "supra", "local")):
            return False
    name = _norm(row.get("issuer", "") or "")
    if name and name == norm_country(row.get("country", "")):
        return True
    return any(h in name for h in SOVEREIGN_NAME_HINTS)


# =============================================================================
# Bond maths (semi-annual, dirty prices, per 100)
# =============================================================================
def cashflow_times(t):
    times, x = [], t
    while x > 1e-9:
        times.append(x)
        x -= 0.5
    return sorted(times)


def bond_price(coupon, t, y):
    """coupon % p.a., t years, y yield as decimal. Dirty price per 100."""
    if t <= 1e-9:
        return 100.0
    c = coupon / 2.0
    df = lambda s: (1 + y / 2.0) ** (-2 * s)
    times = cashflow_times(t)
    return sum(c * df(s) for s in times) + 100.0 * df(times[-1])


def accrued(coupon, t):
    frac_elapsed = 0.5 - (t % 0.5 if t % 0.5 > 1e-9 else 0.5)
    return coupon / 2.0 * frac_elapsed / 0.5


def yield_from_price(coupon, t, clean_price):
    target = clean_price + accrued(coupon, t)
    lo, hi = -0.05, 3.0
    for _ in range(200):
        mid = (lo + hi) / 2
        if bond_price(coupon, t, mid) > target:
            lo = mid
        else:
            hi = mid
    return (lo + hi) / 2


def mod_duration(coupon, t, y):
    h = 0.0001
    p0 = bond_price(coupon, t, y)
    return (bond_price(coupon, t, y - h) - bond_price(coupon, t, y + h)) / (2 * p0 * h)


def horizon_return(coupon, t, y0, y_end, H):
    """Holding-period return over H years: (P_H + coupons received - P_0) / P_0."""
    p0 = bond_price(coupon, t, y0)
    coupons = sum(coupon / 2.0 for s in cashflow_times(t) if s <= H + 1e-9)
    t_end = t - H
    p1 = 100.0 if t_end <= 1e-9 else bond_price(coupon, t_end, y_end)
    return (p1 + coupons - p0) / p0


# =============================================================================
# Curve fitting
# =============================================================================
def fit_curve(ts, ys, min_bonds, sigma):
    """Returns (callable fitted(t)->yield decimal, description) or (None, reason)."""
    ts, ys = np.asarray(ts, float), np.asarray(ys, float)
    if len(ts) < 3:
        return None, f"no curve ({len(ts)} bonds)"

    def design(t, full):
        cols = [np.ones_like(t), np.log1p(t)]
        if full:
            cols.append(t)
        return np.column_stack(cols)

    full = len(ts) >= min_bonds
    X = design(ts, full)
    beta, *_ = np.linalg.lstsq(X, ys, rcond=None)
    resid = ys - X @ beta
    sd = resid.std(ddof=0)
    mask = np.abs(resid) <= sigma * sd if sd > 0 else np.ones_like(ts, bool)
    dropped = int((~mask).sum())
    if dropped and mask.sum() >= (min_bonds if full else 3):
        beta, *_ = np.linalg.lstsq(design(ts[mask], full), ys[mask], rcond=None)
    form = "a + b·ln(1+t) + c·t" if full else "a + b·ln(1+t)"
    t_min, t_max = ts.min(), ts.max()

    def f(t):
        t = min(max(t, 0.0), t_max + 5)            # no wild extrapolation below 0
        v = beta[0] + beta[1] * math.log1p(t) + (beta[2] * t if full else 0.0)
        return v

    desc = f"{form}; n={len(ts)}; outliers dropped={dropped}; range {t_min:.1f}-{t_max:.1f}y"
    return f, desc


# =============================================================================
# Engine
# =============================================================================
def prepare(df, cfg):
    as_of = to_date(cfg["AS_OF"]) if cfg.get("AS_OF") else date.today()
    rows, dropped = [], []
    for _, r in df.iterrows():
        r = r.to_dict()
        isin = str(r.get("isin", "") or "").strip()
        if not re.match(r"^[A-Z]{2}[A-Z0-9]{9}\d$", isin.upper()):
            continue                                  # section/header/blank rows
        rec = {"isin": isin.upper(), "issuer": str(r.get("issuer", "")).strip(),
               "country": str(r.get("country", "")).strip(),
               "currency": str(r.get("currency", "USD") or "USD").strip().upper() if "currency" in df else "USD",
               "view": str(r.get("view", "") or "").strip() if "view" in df else "",
               "seniority": str(r.get("seniority", "") or "").strip() if "seniority" in df else ""}
        if rec["view"].lower() == "nan":
            rec["view"] = ""
        if rec["seniority"].lower() == "nan":
            rec["seniority"] = ""
        coupon, mat = to_float(r.get("coupon")), to_date(r.get("maturity"))
        price, yld = to_float(r.get("price")), to_float(r.get("yield"))
        sp, mo = notch_sp(r.get("sp")), notch_moody(r.get("moody"))
        fi = notch_sp(r.get("fitch")) if "fitch" in df else None
        why = None
        if mat is None:
            why = "perpetual / no maturity"
        elif coupon is None:
            why = "missing coupon (floater?)"
        else:
            t = (mat - as_of).days / 365.25
            if t < cfg["MIN_YEARS_TO_MATURITY"]:
                why = f"matures inside horizon ({t:.2f}y)"
        n = composite_notch(sp, mo, fi)
        if why is None and n is None:
            why = "no rating"
        if why is None and yld is None and price is None:
            why = "no price or yield"
        if why:
            dropped.append({**rec, "reason": why})
            continue
        if yld is not None and abs(yld) < 1.0 and yld != 0:     # given as decimal
            yld *= 100
        if yld is None:
            yld = yield_from_price(coupon, t, price) * 100
        if price is None:
            price = bond_price(coupon, t, yld / 100) - accrued(coupon, t)
        amt = to_float(r.get("amount")) if "amount" in df else None
        if amt is not None and amt > 1e5:
            amt /= 1e6
        rec.update(coupon=coupon, maturity=mat, years=t, price=price, yield_pct=yld,
                   notch=n, rating=notch_to_label(n), sp=r.get("sp"), moody=r.get("moody"),
                   grade=grade_bucket(n), amount_mm=amt,
                   call_date=to_date(r.get("call_date")) if "call_date" in df else None,
                   sovereign=is_sovereign(r), country_key=norm_country(rec["country"]))
        rows.append(rec)
    u = pd.DataFrame(rows)
    if cfg["CURRENCY"] != "all" and len(u):
        off = u[u.currency != cfg["CURRENCY"].upper()]
        for _, x in off.iterrows():
            dropped.append({"isin": x["isin"], "issuer": x.issuer, "country": x.country,
                            "reason": f"currency {x.currency}"})
        u = u[u.currency == cfg["CURRENCY"].upper()].reset_index(drop=True)
    return u, pd.DataFrame(dropped), as_of


def score_universe(u, cfg):
    """Fits curves and computes 12m total returns (base / rally / sell-off) for EVERY bond,
    so peers are measured exactly like the bonds we're choosing."""
    H, share = cfg["HORIZON_YEARS"], cfg["REVERSION_SHARE"]
    curves, curve_info = {}, []
    u = u.copy()
    u["curve_key"] = u.issuer + " | " + u.currency

    # 1) issuer curves; 2) fallback country curve (same ccy, same sov/non-sov flag)
    for key, g in u.groupby("curve_key"):
        f, desc = fit_curve(g.years, g.yield_pct / 100, cfg["CURVE_MIN_BONDS"], cfg["CURVE_OUTLIER_SIGMA"])
        curves[key] = f
        curve_info.append({"curve": key, "fit": desc, "bonds": len(g)})
    for (ck, ccy), g in u.groupby(["country_key", "currency"]):
        key = f"[country] {ck} | {ccy}"
        f, desc = fit_curve(g.years, g.yield_pct / 100, cfg["CURVE_MIN_BONDS"], cfg["CURVE_OUTLIER_SIGMA"])
        curves[key] = f
        curve_info.append({"curve": key, "fit": desc, "bonds": len(g)})

    out = []
    for _, b in u.iterrows():
        f = curves.get(b.curve_key)
        curve_used = b.curve_key
        if f is None:
            curve_used = f"[country] {b.country_key} | {b.currency}"
            f = curves.get(curve_used)
        y0, t, c = b.yield_pct / 100, b.years, b.coupon
        if f is None:
            fit_now, fit_then, curve_used = y0, y0, "none (flat)"
        else:
            fit_now, fit_then = f(t), f(max(t - H, 0.0))
        resid = y0 - fit_now
        beta = cfg["SCENARIO_BETA"].get(b.grade, 1.0)
        shift = cfg["SCENARIO_SHIFT_BP"] / 1e4 * beta
        el = (ANNUAL_PD_BY_NOTCH.get(int(b.notch), 0) / 100 * (1 - cfg["RECOVERY_RATE"]) * H
              if cfg["USE_EXPECTED_LOSS"] else 0.0)

        # static-yield carry, roll-down along the fitted curve, cheapness reversion
        tr_carry = horizon_return(c, t, y0, y0, H)
        y_roll = fit_then + resid                       # rolls down curve, keeps its residual
        tr_roll = horizon_return(c, t, y0, y_roll, H)
        y_base = fit_then + resid * (1 - share)
        tr_base_gross = horizon_return(c, t, y0, y_base, H)
        tr_rally = horizon_return(c, t, y0, max(y_base - shift, -0.01), H) - el
        tr_sell = horizon_return(c, t, y0, y_base + shift, H) - el

        call_flag = ""
        if b.call_date and b.call_date < b.maturity:
            call_flag = "callable" + (" (price > 100: yield may be to call)" if b.price > 100 else "")
        out.append({
            **b.to_dict(),
            "curve_used": curve_used,
            "mod_dur": mod_duration(c, t, y0),
            "fitted_yield_pct": fit_now * 100,
            "vs_curve_bp": resid * 1e4,
            "carry": tr_carry,
            "rolldown": tr_roll - tr_carry,
            "cheapness": tr_base_gross - tr_roll,
            "exp_loss": -(ANNUAL_PD_BY_NOTCH.get(int(b.notch), 0) / 100 * (1 - cfg["RECOVERY_RATE"]) * H),
            "tr_base": tr_base_gross - el,
            "tr_rally": tr_rally,
            "tr_selloff": tr_sell,
            "beta": beta,
            "call_flag": call_flag,
            "subordinated": bool(re.search(r"(?i)sub|junior|hybrid|tier|at1|t2", b.seniority or "")),
        })
    return pd.DataFrame(out), pd.DataFrame(curve_info)


def find_peers(bond, scored, cfg):
    """Tight band first; widen maturity, then rating, only if too few peers. Returns (peers, mat_band, notch_band)."""
    pool = scored[(scored.currency == bond.currency) & (scored.issuer != bond.issuer)]
    if cfg["PEER_EXCLUDE_SAME_COUNTRY"]:
        pool = pool[pool.country_key != bond.country_key]
    if cfg["PEER_EXCLUDE_SELL_VIEW"]:
        pool = pool[~pool.view.str.lower().isin([v.lower() for v in cfg["EXCLUDE_VIEWS"]])]
    nb, nw = cfg["PEER_NOTCH_BAND"], cfg["PEER_NOTCH_BAND_WIDE"]
    mb, mw = cfg["PEER_MATURITY_BAND"], cfg["PEER_MATURITY_BAND_WIDE"]
    steps = [(nb, mb), (nb, mw)] + [(n, mw) for n in range(nb + 1, nw + 1)]
    for notch_band, mat_band in steps:
        peers = pool[((pool.notch - bond.notch).abs() <= notch_band) &
                     ((pool.years - bond.years).abs() <= mat_band)]
        if len(peers) >= cfg["PEER_MIN_COUNT"]:
            break
    peers = peers.assign(distance=(peers.notch - bond.notch).abs() + (peers.years - bond.years).abs() / 2)
    return peers.sort_values("distance"), mat_band, notch_band


def benchmark(scored, targets, cfg):
    rows, peer_tables = [], {}
    for _, b in targets.iterrows():
        peers, band, nband = find_peers(b, scored, cfg)
        n = len(peers)
        rec = b.to_dict()
        rec.update(peer_count=n, peer_band_y=band, peer_band_notch=nband)
        if n >= cfg["PEER_MIN_COUNT"]:
            for k in ("tr_base", "tr_rally", "tr_selloff", "carry", "rolldown", "cheapness", "exp_loss", "yield_pct"):
                rec[f"peer_{k}"] = peers[k].median()
            rec["peer_rating_mix"] = ", ".join(f"{k} {v}" for k, v in peers.rating.value_counts().items())
            rec["beats_pct_of_peers"] = (peers.tr_base < b.tr_base).mean()
            rec["peer_top_quartile"] = peers.tr_base.quantile(0.75)
        else:
            for k in ("tr_base", "tr_rally", "tr_selloff", "carry", "rolldown", "cheapness", "exp_loss", "yield_pct"):
                rec[f"peer_{k}"] = np.nan
            rec["peer_rating_mix"] = ""
            rec["beats_pct_of_peers"] = np.nan
            rec["peer_top_quartile"] = np.nan
        rows.append(rec)
        peer_tables[b["isin"]] = peers
    res = pd.DataFrame(rows)
    for k in ("base", "rally", "selloff"):
        res[f"xs_{k}_bp"] = (res[f"tr_{k}"] - res[f"peer_tr_{k}"]) * 1e4
    w = cfg["SCORE_WEIGHTS"]
    res["penalty_bp"] = 0.0
    if "amount_mm" in res:
        res.loc[res.amount_mm.notna() & (res.amount_mm < cfg["SMALL_ISSUE_MM"]), "penalty_bp"] += cfg["SMALL_ISSUE_PENALTY_BP"]
    res.loc[res.subordinated, "penalty_bp"] += cfg["SUBORDINATED_PENALTY_BP"]
    res["score_bp"] = (w["base"] * res.xs_base_bp + w["rally"] * res.xs_rally_bp
                       + w["selloff"] * res.xs_selloff_bp) / sum(w.values()) - res.penalty_bp
    res["robust"] = (res.xs_base_bp > 0) & (res.xs_rally_bp > 0) & (res.xs_selloff_bp > 0)
    res["bucket"] = res.years.apply(lambda t: next((n for n, lo, hi in cfg["BUCKETS"] if lo <= t < hi), "Long"))
    return res, peer_tables


def pct(x, d=1):
    return "n/a" if x is None or pd.isna(x) else f"{x * 100:+.{d}f}%"


def explain(b, cfg):
    if pd.isna(b.peer_tr_base):
        return (f"Only {int(b.peer_count)} similarly rated peers (need {cfg['PEER_MIN_COUNT']}); "
                f"cannot be benchmarked. Expected 12m TR {pct(b.tr_base)}.")
    parts = [
        f"Expected 12m total return {b.tr_base * 100:.1f}% vs {b.peer_tr_base * 100:.1f}% median for "
        f"{int(b.peer_count)} peers rated {b.rating}±{int(b.peer_band_notch)} notch maturing within ±{b.peer_band_y:g}y "
        f"→ {b.xs_base_bp:+.0f}bp; beats {b.beats_pct_of_peers * 100:.0f}% of the peer group."
    ]
    drivers = {
        "carry": (b.carry - b.peer_carry) * 1e4,
        "roll-down": (b.rolldown - b.peer_rolldown) * 1e4,
        "curve cheapness": (b.cheapness - b.peer_cheapness) * 1e4,
        **({"expected loss": (b.exp_loss - b.peer_exp_loss) * 1e4} if cfg["USE_EXPECTED_LOSS"] else {}),
    }
    ranked = sorted(drivers.items(), key=lambda kv: -abs(kv[1]))
    parts.append("Edge vs peers comes from " + ", ".join(f"{k} {v:+.0f}bp" for k, v in ranked if abs(v) >= 5) + "."
                 if any(abs(v) >= 5 for v in drivers.values()) else "Edge vs peers is small on every driver.")
    if abs(b.vs_curve_bp) >= 10:
        parts.append(f"Trades {abs(b.vs_curve_bp):.0f}bp {'cheap' if b.vs_curve_bp > 0 else 'rich'} "
                     f"to its fitted curve ({b.curve_used.split(' | ')[0]}).")
    parts.append(f"Rally (-{cfg['SCENARIO_SHIFT_BP'] * b.beta:.0f}bp): {b.xs_rally_bp:+.0f}bp vs peers; "
                 f"sell-off (+{cfg['SCENARIO_SHIFT_BP'] * b.beta:.0f}bp): {b.xs_selloff_bp:+.0f}bp vs peers."
                 + (" Outperforms in all three scenarios." if b.robust else ""))
    cav = []
    if b.call_flag:
        cav.append(b.call_flag)
    if b.subordinated:
        cav.append("subordinated")
    if b.amount_mm is not None and not pd.isna(b.amount_mm) and b.amount_mm < cfg["SMALL_ISSUE_MM"]:
        cav.append(f"small issue (${b.amount_mm:,.0f}mm)")
    if "none" in b.curve_used or "[country]" in b.curve_used:
        cav.append(f"issuer curve too thin — used {b.curve_used}")
    if b.view:
        cav.append(f"CIO view: {b.view}")
    if cav:
        parts.append("Watch: " + "; ".join(cav) + ".")
    return " ".join(parts)


def select_picks(res, cfg):
    excl = [v.lower() for v in cfg["EXCLUDE_VIEWS"]]
    eligible = res[~res.view.str.lower().isin(excl) & res.peer_tr_base.notna()]
    picks = []
    for curve, g in eligible.groupby("curve_key"):
        g = g.sort_values("score_bp", ascending=False)
        winners = g[g.xs_base_bp > cfg["MIN_EXCESS_BP"]]
        for rank, (_, b) in enumerate(winners.head(cfg["TOP_N_PER_CURVE"]).iterrows(), 1):
            picks.append({**b.to_dict(), "pick_type": f"Top {rank} on curve"})
        for name, lo, hi in cfg["BUCKETS"]:
            gb = g[g.bucket == name]
            if gb.empty:
                continue
            best = gb.iloc[0]
            label = f"Best {name}" + ("" if best.xs_base_bp > cfg["MIN_EXCESS_BP"] else " — does NOT beat peers")
            picks.append({**best.to_dict(), "pick_type": label})
    p = pd.DataFrame(picks)
    if len(p):                                   # one row per bond; merge labels if it wins several ways
        labels = p.groupby(["curve_key", "isin"], sort=False).pick_type.apply(lambda x: " + ".join(x))
        p = p.drop_duplicates(["curve_key", "isin"]).set_index(["curve_key", "isin"])
        p["pick_type"] = labels
        p = p.reset_index()
        p["why"] = p.apply(lambda b: explain(b, cfg), axis=1)
    return p


# =============================================================================
# Excel writer
# =============================================================================
def write_excel(path, cfg, results, picks, peer_tables, curve_info, dropped, as_of, src, mapping, title):
    from openpyxl import Workbook
    from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
    from openpyxl.utils import get_column_letter

    RED, GREY, LIGHT = "E60000", "404040", "F2F2F2"
    hfont = Font(name="Arial", bold=True, color="FFFFFF", size=9)
    hfill = PatternFill("solid", fgColor=GREY)
    body = Font(name="Arial", size=9)
    bold = Font(name="Arial", size=9, bold=True)
    thin = Border(bottom=Side(style="thin", color="D9D9D9"))
    good = PatternFill("solid", fgColor="E2EFDA")
    bad = PatternFill("solid", fgColor="FCE4D6")
    wb = Workbook()

    def sheet_title(ws, text, sub):
        ws["A1"] = text; ws["A1"].font = Font(name="Arial", size=14, bold=True, color=RED)
        ws["A2"] = sub; ws["A2"].font = Font(name="Arial", size=9, italic=True, color=GREY)

    def table(ws, start_row, cols, rows, widths=None, xs_formula=None):
        """cols: list of (header, key, fmt). xs_formula: {col_header: (tr_header, peer_header)}"""
        for j, (h, _, _) in enumerate(cols, 1):
            c = ws.cell(row=start_row, column=j, value=h)
            c.font, c.fill = hfont, hfill
            c.alignment = Alignment(wrap_text=True, vertical="center", horizontal="center")
        ws.row_dimensions[start_row].height = 30
        pos = {h: get_column_letter(j) for j, (h, _, _) in enumerate(cols, 1)}
        for i, r in enumerate(rows, start_row + 1):
            for j, (h, key, fmt) in enumerate(cols, 1):
                if xs_formula and h in xs_formula:
                    a, bcol = xs_formula[h]
                    v = f'=IF(OR({pos[a]}{i}="",{pos[bcol]}{i}=""),"",({pos[a]}{i}-{pos[bcol]}{i})*10000)'
                else:
                    v = r.get(key) if isinstance(r, dict) else r[key]
                    if isinstance(v, float) and math.isnan(v):
                        v = None
                    if isinstance(v, (np.bool_, bool)):
                        v = "Yes" if v else ""
                    if isinstance(v, np.generic):
                        v = v.item()
                c = ws.cell(row=i, column=j, value=v)
                c.font, c.border = body, thin
                if fmt:
                    c.number_format = fmt
                if h == "Why":
                    c.alignment = Alignment(wrap_text=True, vertical="top")
        for j, (h, _, _) in enumerate(cols, 1):
            ws.column_dimensions[get_column_letter(j)].width = (widths or {}).get(h, max(9, min(22, len(h) + 2)))
        ws.freeze_panes = ws.cell(row=start_row + 1, column=4)
        return pos

    P, P2, BP, D = "0.00%", "0.00", '+0;-0;0', "yyyy-mm-dd"
    core = [("ISIN", "isin", None), ("Issuer", "issuer", None), ("Country", "country", None),
            ("Coupon", "coupon", "0.000"), ("Maturity", "maturity", D), ("Years", "years", "0.0"),
            ("Rating", "rating", None), ("S&P", "sp", None), ("Moody's", "moody", None),
            ("Price", "price", "0.00"), ("Yield %", "yield_pct", P2), ("Mod dur", "mod_dur", "0.0"),
            ("Fitted yld %", "fitted_yield_pct", P2), ("vs curve bp", "vs_curve_bp", BP),
            ("Carry", "carry", P), ("Roll-down", "rolldown", P), ("Cheapness", "cheapness", P),
            ("Exp. loss (info)", "exp_loss", P), ("12m TR base", "tr_base", P),
            ("Peer median TR", "peer_tr_base", P), ("Excess vs peers bp", None, BP),
            ("Peers (n)", "peer_count", "0"), ("Peer notch ±", "peer_band_notch", "0"), ("Beats % peers", "beats_pct_of_peers", "0%"),
            ("TR rally", "tr_rally", P), ("Peer TR rally", "peer_tr_rally", P), ("Excess rally bp", None, BP),
            ("TR sell-off", "tr_selloff", P), ("Peer TR sell-off", "peer_tr_selloff", P),
            ("Excess sell-off bp", None, BP), ("Penalty bp", "penalty_bp", "0"),
            ("Score bp", "score_bp", BP), ("All 3 scenarios", "robust", None),
            ("Bucket", "bucket", None), ("CIO view", "view", None), ("Flags", "call_flag", None)]
    xs = {"Excess vs peers bp": ("12m TR base", "Peer median TR"),
          "Excess rally bp": ("TR rally", "Peer TR rally"),
          "Excess sell-off bp": ("TR sell-off", "Peer TR sell-off")}

    # ---- Summary ----------------------------------------------------------
    ws = wb.active; ws.title = "Summary"
    sheet_title(ws, title, f"Best bonds per curve on 12m total return vs similarly rated peers | "
                           f"as of {as_of} | source: {os.path.basename(src)} | INDICATIVE — model output, not a recommendation")
    srows = []
    if len(picks):
        for _, p in picks[picks.pick_type.str.startswith("Top 1 on curve")].iterrows():
            srows.append(p.to_dict())
    scols = [("Curve", "curve_key", None), ("ISIN", "isin", None), ("Coupon", "coupon", "0.000"),
             ("Maturity", "maturity", D), ("Rating", "rating", None), ("Yield %", "yield_pct", P2),
             ("12m TR base", "tr_base", P), ("Peer median TR", "peer_tr_base", P),
             ("Excess vs peers bp", None, BP), ("Beats % peers", "beats_pct_of_peers", "0%"),
             ("All 3 scenarios", "robust", None), ("Why", "why", None)]
    table(ws, 4, scols, srows, widths={"Curve": 30, "Why": 110},
          xs_formula={"Excess vs peers bp": ("12m TR base", "Peer median TR")})
    for i in range(5, 5 + len(srows)):
        ws.row_dimensions[i].height = 75
    no_pick = sorted(set(results.curve_key) - set(r["curve_key"] for r in srows))
    if no_pick:
        r0 = 6 + len(srows)
        ws.cell(row=r0, column=1, value="Curves with no bond beating its peer median:").font = bold
        for k, name in enumerate(no_pick, 1):
            ws.cell(row=r0 + k, column=1, value=name).font = body

    # ---- Best picks -------------------------------------------------------
    ws = wb.create_sheet("Best Picks")
    sheet_title(ws, "Best picks per curve and per maturity bucket", "Why = generated explanation. Excess bp columns are live formulas.")
    pcols = [("Pick", "pick_type", None), ("Curve", "curve_key", None)] + core[:1] + core[3:8] + \
            [("Yield %", "yield_pct", P2), ("12m TR base", "tr_base", P), ("Peer median TR", "peer_tr_base", P),
             ("Excess vs peers bp", None, BP), ("TR rally", "tr_rally", P), ("Peer TR rally", "peer_tr_rally", P),
             ("Excess rally bp", None, BP), ("TR sell-off", "tr_selloff", P),
             ("Peer TR sell-off", "peer_tr_selloff", P), ("Excess sell-off bp", None, BP),
             ("Score bp", "score_bp", BP), ("Why", "why", None)]
    prow = picks.to_dict("records") if len(picks) else []
    table(ws, 4, pcols, prow, widths={"Pick": 34, "Curve": 28, "Why": 100}, xs_formula=xs)
    for i, r in enumerate(prow, 5):
        ws.row_dimensions[i].height = 80
        if "NOT" in r["pick_type"]:
            ws.cell(row=i, column=1).fill = bad
        elif r["pick_type"].startswith("Top 1 on curve"):
            ws.cell(row=i, column=1).fill = good

    # ---- one tab per country / curve group --------------------------------
    for country, g in results.groupby("country"):
        name = re.sub(r"[\[\]:*?/\\]", "", str(country))[:31] or "Other"
        if name in wb.sheetnames:
            name = name[:28] + "_2"
        ws = wb.create_sheet(name)
        sheet_title(ws, f"{country} — full curve, ranked by score",
                    "Green = beats peer median in all three scenarios; orange = trails peers in the base case.")
        g = g.sort_values(["curve_key", "score_bp"], ascending=[True, False])
        rows = g.to_dict("records")
        table(ws, 4, [("Curve", "curve_key", None)] + core, rows, widths={"Curve": 28}, xs_formula=xs)
        for i, r in enumerate(rows, 5):
            fill = good if r["robust"] else (bad if (r["xs_base_bp"] is not None and not pd.isna(r["xs_base_bp"]) and r["xs_base_bp"] < 0) else None)
            if fill:
                ws.cell(row=i, column=2).fill = fill

    # ---- peer detail ------------------------------------------------------
    ws = wb.create_sheet("Peer Detail")
    sheet_title(ws, "Peer groups behind each Top-1 pick", "Median uses ALL matching peers; closest 15 shown.")
    r0 = 4
    top1 = picks[picks.pick_type.str.startswith("Top 1 on curve")] if len(picks) else picks
    peer_cols = [("ISIN", "isin", None), ("Issuer", "issuer", None), ("Country", "country", None),
                 ("Coupon", "coupon", "0.000"), ("Maturity", "maturity", D), ("Years", "years", "0.0"),
                 ("Rating", "rating", None), ("Yield %", "yield_pct", P2), ("12m TR base", "tr_base", P),
                 ("TR rally", "tr_rally", P), ("TR sell-off", "tr_selloff", P)]
    for _, p in top1.iterrows():
        ws.cell(row=r0, column=1, value=f"{p.issuer} {p.coupon:g}% {p.maturity} ({p["isin"]}) — {p.rating}, "
                                        f"TR {p.tr_base * 100:.2f}% vs peer median {p.peer_tr_base * 100:.2f}%").font = bold
        rows = peer_tables[p["isin"]].head(15).to_dict("records")
        for j, (h, _, _) in enumerate(peer_cols, 1):
            c = ws.cell(row=r0 + 1, column=j, value=h); c.font, c.fill = hfont, hfill
        for i, r in enumerate(rows, r0 + 2):
            for j, (h, key, fmt) in enumerate(peer_cols, 1):
                v = r.get(key)
                c = ws.cell(row=i, column=j, value=v.item() if isinstance(v, np.generic) else v)
                c.font = body
                if fmt:
                    c.number_format = fmt
        r0 += len(rows) + 3
    for j, w in enumerate([14, 30, 14, 8, 11, 7, 7, 8, 11, 10, 10], 1):
        ws.column_dimensions[get_column_letter(j)].width = w

    # ---- curves / dropped / methodology ------------------------------------
    ws = wb.create_sheet("Curve Fits")
    sheet_title(ws, "Fitted curves", "Issuer curves first; [country] curves are the fallback for thin issuers.")
    table(ws, 4, [("Curve", "curve", None), ("Bonds", "bonds", "0"), ("Fit", "fit", None)],
          curve_info.to_dict("records"), widths={"Curve": 40, "Fit": 70})

    ws = wb.create_sheet("Dropped")
    sheet_title(ws, "Bonds excluded from the analysis", "Check here first if an ISIN you expected is missing.")
    table(ws, 4, [("ISIN", "isin", None), ("Issuer", "issuer", None), ("Country", "country", None),
                  ("Reason", "reason", None)], dropped.to_dict("records") if len(dropped) else [],
          widths={"Issuer": 30, "Reason": 40})

    ws = wb.create_sheet("Methodology")
    sheet_title(ws, "Methodology & settings", "Every number below is read from CONFIG at run time.")
    lines = [
        ("KPI", "Expected 12-month total return of the bond minus the MEDIAN expected return of similarly rated peers, same model for both."),
        ("Total return", "(Price in 12m + coupons received − price today) / price today, full semi-annual cash-flow repricing (no duration shortcut), less expected credit loss."),
        ("Carry", "Return if the bond's yield is unchanged in 12m."),
        ("Roll-down", "Extra return from rolling down its own fitted curve: yield in 12m = fitted yield at (maturity − 1y) + today's residual."),
        ("Curve cheapness", f"{cfg['REVERSION_SHARE']:.0%} of the bond's gap to its fitted curve assumed to close over the horizon."),
        ("Curve fit", f"Least squares a + b·ln(1+t) + c·t per issuer (≥{cfg['CURVE_MIN_BONDS']} bonds; 3 bonds drop c), refit after removing >{cfg['CURVE_OUTLIER_SIGMA']}σ outliers; thin issuers use a same-country curve."),
        ("Expected loss", f"{'ON' if cfg['USE_EXPECTED_LOSS'] else 'OFF (shown for information only, not deducted)'} — approx. long-run 1y default probability for the composite rating × (1 − {cfg['RECOVERY_RATE']:.0%} recovery). Composite = middle of three agencies, lower of two."),
        ("Peers", f"Whole universe, same currency, different issuer{', different country' if cfg['PEER_EXCLUDE_SAME_COUNTRY'] else ''}, composite rating ±{cfg['PEER_NOTCH_BAND']} notch, maturity ±{cfg['PEER_MATURITY_BAND']}y If fewer than {cfg['PEER_MIN_COUNT']}: maturity widened to ±{cfg['PEER_MATURITY_BAND_WIDE']}y, then rating one notch at a time up to ±{cfg['PEER_NOTCH_BAND_WIDE']} (band used is stated in each Why). Still too few → not benchmarked."),
        ("Scenarios", f"Parallel ±{cfg['SCENARIO_SHIFT_BP']}bp × rating beta {cfg['SCENARIO_BETA']}, applied to the bond AND its peers."),
        ("Score", f"Weighted excess vs peers {cfg['SCORE_WEIGHTS']} minus penalties (issue < ${cfg['SMALL_ISSUE_MM']}mm: {cfg['SMALL_ISSUE_PENALTY_BP']}bp; subordinated: {cfg['SUBORDINATED_PENALTY_BP']}bp)."),
        ("Picks", f"Top {cfg['TOP_N_PER_CURVE']} per curve with base excess > {cfg['MIN_EXCESS_BP']}bp, plus the best bond in each bucket {[b[0] for b in cfg['BUCKETS']]}. CIO views {cfg['EXCLUDE_VIEWS']} never picked."),
        ("Known limits", "Yield-based (no UST curve, so spread = yield vs curve); no liquidity/bid-ask data beyond issue size; callables priced to maturity unless the feed yield is YTW; indicative prices."),
        ("Universe", f"{cfg['UNIVERSE']} | currency {cfg['CURRENCY']} | as of {as_of}"),
        ("Input", f"{src}"),
        ("Columns", ", ".join(f"{k}←{v}" for k, v in mapping.items())),
    ]
    for i, (k, v) in enumerate(lines, 4):
        ws.cell(row=i, column=1, value=k).font = bold
        c = ws.cell(row=i, column=2, value=v); c.font = body
        c.alignment = Alignment(wrap_text=True, vertical="top")
    ws.column_dimensions["A"].width = 18
    ws.column_dimensions["B"].width = 130

    os.makedirs(os.path.dirname(path) or ".", exist_ok=True)
    wb.save(path)
    return path


# =============================================================================
# Orchestration (used by both scripts)
# =============================================================================
def run(cfg, select, label, verbose=True):
    """select(prepared_universe) -> boolean mask of the bonds to choose from."""
    raw, src, mapping = load_universe(cfg, verbose)
    u, dropped, as_of = prepare(raw, cfg)
    if u.empty:
        raise SystemExit("No usable bonds after cleaning — see column detection above.")
    scored, curve_info = score_universe(u, cfg)
    mask = select(scored)
    targets = scored[mask]
    if targets.empty:
        raise SystemExit(f"No bonds matched {label}.")
    results, peer_tables = benchmark(scored, targets, cfg)
    picks = select_picks(results, cfg)
    keys = set(targets.curve_key)
    curve_info = curve_info[curve_info.curve.isin(keys) |
                            curve_info.curve.str.startswith("[country]") &
                            curve_info.curve.str.contains("|".join(re.escape(c) for c in set(targets.country_key)))]
    tgt_isins = set(targets["isin"])
    dropped_rel = dropped[dropped.country.map(norm_country).isin(set(targets.country_key))] if len(dropped) else dropped
    name = cfg.get("OUTPUT_NAME") or f"Best_Bonds_{label.replace(' ', '_')}_{as_of:%Y%m%d}.xlsx"
    out = os.path.join(cfg["OUTPUT_DIR"], name)
    write_excel(out, cfg, results, picks, peer_tables, curve_info, dropped_rel, as_of, src, mapping,
                f"Best bonds vs peers — {label}")
    if verbose:
        print(f"\nUniverse scored: {len(scored)} bonds | chosen from: {len(tgt_isins)} | peers from whole universe")
        top1 = picks[picks.pick_type.str.startswith("Top 1 on curve")] if len(picks) else picks
        for _, p in top1.iterrows():
            print(f"  {p.curve_key:<38} {p["isin"]}  {p.coupon:g}% {p.maturity}  TR {p.tr_base * 100:5.2f}% "
                  f"vs peers {p.peer_tr_base * 100:5.2f}%  ({p.xs_base_bp:+.0f}bp)")
        print(f"\nSaved: {out}")
    return out, results, picks


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--country", action="append")
    ap.add_argument("--data")
    ap.add_argument("--universe", choices=["sovereign", "all"])
    ap.add_argument("--ccy")
    ap.add_argument("--as-of")
    ap.add_argument("--out")
    ap.add_argument("--show-columns", action="store_true")
    a = ap.parse_args()
    cfg = dict(CONFIG)
    if a.country: cfg["COUNTRIES"] = a.country
    if a.data: cfg["DATA_PATH"] = a.data
    if a.universe: cfg["UNIVERSE"] = a.universe
    if a.ccy: cfg["CURRENCY"] = a.ccy
    if a.as_of: cfg["AS_OF"] = a.as_of
    if a.out: cfg["OUTPUT_NAME"] = a.out
    if a.show_columns:
        load_universe(cfg); return

    wanted = {norm_country(c) for c in cfg["COUNTRIES"]}

    def select(s):
        m = s.country_key.isin(wanted)
        if cfg["UNIVERSE"] == "sovereign":
            m &= s.sovereign
        return m

    found = set()
    out, results, _ = run(cfg, select, "-".join(cfg["COUNTRIES"]))
    found = set(results.country_key)
    for c in wanted - found:
        print(f"  ! no {cfg['UNIVERSE']} bonds found for '{c}' — check spelling / --universe all")


if __name__ == "__main__":
    main()
