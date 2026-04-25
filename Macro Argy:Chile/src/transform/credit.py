"""Credit-specific helpers: USD curve assembly, brecha (FX gap) calc."""

from __future__ import annotations

from typing import List, Tuple

import pandas as pd

from src.config_loader import CountryConfig


def build_curve(
    cfg: CountryConfig,
    prices: pd.DataFrame,
) -> pd.DataFrame:
    """Construct the latest sovereign USD curve as a (maturity_years, yield) table.

    `prices` is the wide DataFrame from CacheStore with bond yield series as columns
    named by series.name. We pull the most recent observation for each bond in the
    `usd_curve` category.
    """
    points: List[Tuple[float, float, str]] = []
    for s in cfg.series:
        if s.category != "usd_curve" or s.curve_maturity_years is None:
            continue
        if s.name not in prices.columns:
            continue
        last = prices[s.name].dropna()
        if last.empty:
            continue
        points.append((s.curve_maturity_years, float(last.iloc[-1]), s.label))
    if not points:
        return pd.DataFrame(columns=["maturity_years", "yield_pct", "bond"])
    df = pd.DataFrame(points, columns=["maturity_years", "yield_pct", "bond"])
    df = df.sort_values("maturity_years").reset_index(drop=True)
    return df


def curve_history(
    cfg: CountryConfig,
    prices: pd.DataFrame,
    snapshots: int = 4,
) -> pd.DataFrame:
    """Return curve at several historical snapshots (today, -1m, -3m, -1y).
    Useful for waterfall curve charts on the Credit tab.
    """
    bonds = [(s.name, s.curve_maturity_years, s.label)
             for s in cfg.series
             if s.category == "usd_curve" and s.curve_maturity_years is not None
             and s.name in prices.columns]
    if not bonds or prices.empty:
        return pd.DataFrame()

    now = prices.index.max()
    offsets = {
        "Today": pd.DateOffset(days=0),
        "-1m": pd.DateOffset(months=1),
        "-3m": pd.DateOffset(months=3),
        "-1y": pd.DateOffset(years=1),
    }
    rows = []
    for label, off in list(offsets.items())[:snapshots]:
        target = now - off
        # nearest available date <= target
        idx = prices.index[prices.index <= target]
        if len(idx) == 0:
            continue
        ts = idx.max()
        for name, mat, _ in bonds:
            v = prices.loc[ts, name] if name in prices.columns else None
            if pd.notna(v):
                rows.append({"snapshot": label, "date": ts, "maturity_years": mat,
                             "bond": name, "yield_pct": float(v)})
    if not rows:
        return pd.DataFrame()
    return pd.DataFrame(rows).sort_values(["snapshot", "maturity_years"])


def brecha(official: pd.Series, parallel: pd.Series) -> pd.Series:
    """Argentine FX gap: (parallel / official - 1) * 100. Aligns on common index."""
    if official is None or parallel is None or official.empty or parallel.empty:
        return pd.Series(dtype=float)
    df = pd.concat([official.rename("o"), parallel.rename("p")], axis=1).dropna()
    return ((df["p"] / df["o"]) - 1) * 100
