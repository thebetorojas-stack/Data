"""Parquet-backed cache for time series.

Layout: cache/<iso>/<series_name>.parquet — one file per series.
Each file holds a single-column DataFrame indexed by date with metadata
(provider, ticker, frequency, last_pulled_utc) in the parquet metadata.

Refresh logic:
- On each refresh, find the last cached date.
- Pull fresh data starting `always_repull_days` before that date (handles revisions).
- Merge into cache, dedupe on index keeping the latest pull, write back.

This is fast (parquet append-equivalent), survives restarts, and gives the
dashboard instant reads without re-hitting Bloomberg/Haver.
"""

from __future__ import annotations

from dataclasses import dataclass
from datetime import date, datetime, timedelta
from pathlib import Path
from typing import Dict, List, Optional, Tuple

import pandas as pd

from src.config_loader import SeriesConfig
from src.fetchers.base import FetchResult


@dataclass
class CacheStore:
    root: Path
    iso: str

    def __post_init__(self):
        self.dir.mkdir(parents=True, exist_ok=True)

    @property
    def dir(self) -> Path:
        return Path(self.root) / self.iso.lower()

    # ------------------------------------------------------------------ #
    def path_for(self, series_name: str) -> Path:
        return self.dir / f"{series_name}.parquet"

    def load(self, series_name: str) -> Optional[pd.Series]:
        p = self.path_for(series_name)
        if not p.exists():
            return None
        df = pd.read_parquet(p)
        if df.empty:
            return None
        s = df.iloc[:, 0]
        s.name = series_name
        s.index = pd.to_datetime(s.index)
        return s

    def save(self, series_name: str, s: pd.Series) -> None:
        if s is None or s.empty:
            return
        s = s.copy()
        s.name = series_name
        s.index = pd.to_datetime(s.index)
        s = s[~s.index.duplicated(keep="last")].sort_index()
        df = s.to_frame()
        self.path_for(series_name).parent.mkdir(parents=True, exist_ok=True)
        df.to_parquet(self.path_for(series_name), engine="pyarrow")

    def last_date(self, series_name: str) -> Optional[date]:
        s = self.load(series_name)
        if s is None or s.empty:
            return None
        return s.index.max().date()

    # ------------------------------------------------------------------ #
    def merge(self, series_name: str, new: pd.Series) -> pd.Series:
        """Append new observations, prefer new on duplicate dates (handles revisions)."""
        if new is None or new.empty:
            return self.load(series_name) or pd.Series(dtype=float)
        existing = self.load(series_name)
        if existing is None or existing.empty:
            merged = new
        else:
            merged = pd.concat([existing, new])
            merged = merged[~merged.index.duplicated(keep="last")].sort_index()
        self.save(series_name, merged)
        return merged

    # ------------------------------------------------------------------ #
    def plan_pull_window(
        self,
        spec: SeriesConfig,
        default_start: date,
        always_repull_days: int = 30,
    ) -> Tuple[date, date]:
        """Return (start, end) for the upcoming pull.

        - If no cache exists, pull from default_start to today.
        - Else pull from (last_cached - always_repull_days) to today, to capture revisions.
        """
        end = date.today()
        last = self.last_date(spec.name)
        if last is None:
            return default_start, end
        start = last - timedelta(days=always_repull_days)
        if start < default_start:
            start = default_start
        return start, end

    # ------------------------------------------------------------------ #
    def apply_results(self, results: List[FetchResult]) -> Dict[str, str]:
        """Persist fetch results, return per-series status dict."""
        statuses: Dict[str, str] = {}
        for r in results:
            if not r.ok:
                statuses[r.name] = f"error: {r.error}"
                continue
            self.merge(r.name, r.series)
            statuses[r.name] = f"ok ({len(r.series)} obs through {r.series.index.max().date()})"
        return statuses

    # ------------------------------------------------------------------ #
    def load_all(self, names: List[str]) -> pd.DataFrame:
        """Return a wide DataFrame with the requested series as columns."""
        frames = []
        for n in names:
            s = self.load(n)
            if s is not None and not s.empty:
                frames.append(s.rename(n))
        if not frames:
            return pd.DataFrame()
        return pd.concat(frames, axis=1).sort_index()
