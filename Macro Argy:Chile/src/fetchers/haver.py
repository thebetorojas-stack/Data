"""Haver Analytics fetcher.

The official Haver Python API (`Haver` module) ships with the DLX desktop install.
Typical install path on Windows: C:\\DLX\\Data, which must be on PYTHONPATH.

Ticker convention: <code>@<database>, e.g. "c917cpi@chile" — code "c917cpi" in
database "chile". The Python API expects database and code separately for some
calls; we parse the @ in the ticker.
"""

from __future__ import annotations

from datetime import date
from typing import Optional

import pandas as pd

from src.config_loader import SeriesConfig
from .base import Fetcher, FetchResult


class HaverFetcher(Fetcher):
    provider = "haver"

    def __init__(self):
        try:
            import Haver  # type: ignore  # noqa: F401
            self._available = True
        except ImportError:
            self._available = False

    # ------------------------------------------------------------------ #
    def health_check(self) -> bool:
        if not self._available:
            return False
        try:
            import Haver
            # GDP@USECON is a sentinel that exists for any Haver subscriber.
            df = Haver.data(["gdp"], database="USECON",
                            startdate="2024-01-01", enddate="2024-06-30")
            return df is not None and len(df) > 0
        except Exception:
            return False

    # ------------------------------------------------------------------ #
    @staticmethod
    def _parse_ticker(ticker: str) -> tuple[str, str]:
        """Split 'code@db' into (code, db). Default db = 'USECON'."""
        if "@" in ticker:
            code, db = ticker.split("@", 1)
            return code.strip().lower(), db.strip().lower()
        return ticker.strip().lower(), "usecon"

    # ------------------------------------------------------------------ #
    def fetch_one(
        self,
        spec: SeriesConfig,
        start: date,
        end: Optional[date] = None,
    ) -> FetchResult:
        if not self._available:
            return FetchResult(spec.name, pd.Series(dtype=float), ok=False,
                               error="Haver Python module not importable. "
                                     "Verify DLX install and PYTHONPATH.")
        try:
            import Haver

            code, db = self._parse_ticker(spec.ticker)
            end = end or date.today()
            df = Haver.data(
                [code],
                database=db,
                startdate=start.isoformat(),
                enddate=end.isoformat(),
            )
            if df is None or df.empty:
                return FetchResult(spec.name, pd.Series(dtype=float), ok=False,
                                   error="empty Haver response")
            # Haver returns a DataFrame with one column named after the code.
            s = df.iloc[:, 0]
            s.name = spec.name
            s.index = pd.to_datetime(s.index)
            return FetchResult(spec.name, s.dropna(), ok=True)
        except Exception as e:  # noqa: BLE001
            return FetchResult(spec.name, pd.Series(dtype=float), ok=False, error=str(e))
