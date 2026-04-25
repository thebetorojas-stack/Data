"""Bloomberg fetcher built on xbbg / blpapi.

Requirements:
- Bloomberg Terminal running and logged in on this machine.
- Desktop API enabled (default).
- `pip install xbbg blpapi`.

Batching: blp.bdh accepts a list of tickers and returns a multi-column DataFrame.
We exploit that for big speedups when refreshing many series at once.
"""

from __future__ import annotations

from datetime import date, timedelta
from typing import List, Optional

import pandas as pd

from src.config_loader import SeriesConfig
from .base import Fetcher, FetchResult


class BloombergFetcher(Fetcher):
    provider = "bloomberg"

    def __init__(self):
        try:
            from xbbg import blp  # noqa: F401
            self._available = True
        except ImportError:
            self._available = False

    # ------------------------------------------------------------------ #
    def health_check(self) -> bool:
        if not self._available:
            return False
        try:
            from xbbg import blp
            df = blp.bdp("SPX Index", "PX_LAST")
            return df is not None and not df.empty
        except Exception:
            return False

    # ------------------------------------------------------------------ #
    def fetch_one(
        self,
        spec: SeriesConfig,
        start: date,
        end: Optional[date] = None,
    ) -> FetchResult:
        if not self._available:
            return FetchResult(spec.name, pd.Series(dtype=float), ok=False,
                               error="xbbg / blpapi not importable")
        try:
            from xbbg import blp

            end = end or date.today()
            df = blp.bdh(
                tickers=spec.ticker,
                flds=spec.field,
                start_date=start.isoformat(),
                end_date=end.isoformat(),
            )
            if df is None or df.empty:
                return FetchResult(spec.name, pd.Series(dtype=float), ok=False,
                                   error="empty response")
            # blp.bdh returns multi-index columns: (ticker, field). Flatten.
            s = df.iloc[:, 0]
            s.name = spec.name
            s.index = pd.to_datetime(s.index)
            return FetchResult(spec.name, s.dropna(), ok=True)
        except Exception as e:  # noqa: BLE001
            return FetchResult(spec.name, pd.Series(dtype=float), ok=False, error=str(e))

    # ------------------------------------------------------------------ #
    def fetch_many(
        self,
        specs: List[SeriesConfig],
        start: date,
        end: Optional[date] = None,
    ) -> List[FetchResult]:
        """Batched pull: groups specs by field, calls blp.bdh once per field."""
        if not self._available:
            return [FetchResult(s.name, pd.Series(dtype=float), ok=False,
                                error="xbbg unavailable") for s in specs]
        from xbbg import blp

        end = end or date.today()
        results: List[FetchResult] = []

        # group specs by field for batched calls
        by_field: dict[str, List[SeriesConfig]] = {}
        for s in specs:
            by_field.setdefault(s.field, []).append(s)

        for fld, group in by_field.items():
            tickers = [g.ticker for g in group]
            try:
                df = blp.bdh(
                    tickers=tickers,
                    flds=fld,
                    start_date=start.isoformat(),
                    end_date=end.isoformat(),
                )
            except Exception as e:  # noqa: BLE001
                for g in group:
                    results.append(FetchResult(g.name, pd.Series(dtype=float),
                                               ok=False, error=str(e)))
                continue

            if df is None or df.empty:
                for g in group:
                    results.append(FetchResult(g.name, pd.Series(dtype=float),
                                               ok=False, error="empty response"))
                continue

            # df has multi-index columns: (ticker, field). Pull each column.
            df.index = pd.to_datetime(df.index)
            for g in group:
                try:
                    if isinstance(df.columns, pd.MultiIndex):
                        s = df[(g.ticker, fld)]
                    else:
                        s = df[g.ticker] if g.ticker in df.columns else df.iloc[:, 0]
                    s = s.dropna()
                    s.name = g.name
                    results.append(FetchResult(g.name, s, ok=True))
                except KeyError:
                    results.append(FetchResult(g.name, pd.Series(dtype=float),
                                               ok=False, error=f"ticker not in response"))
        return results
