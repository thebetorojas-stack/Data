"""Abstract Fetcher interface and shared types."""

from __future__ import annotations

from abc import ABC, abstractmethod
from dataclasses import dataclass
from datetime import date
from typing import List, Optional

import pandas as pd

from src.config_loader import SeriesConfig


@dataclass
class FetchResult:
    name: str
    series: pd.Series          # date-indexed, name=series.name
    ok: bool
    error: Optional[str] = None


class Fetcher(ABC):
    """Each provider implements fetch_one() returning a FetchResult.
    The cache layer handles batching and delta logic.
    """

    provider: str = ""

    @abstractmethod
    def fetch_one(
        self,
        spec: SeriesConfig,
        start: date,
        end: Optional[date] = None,
    ) -> FetchResult:
        ...

    def fetch_many(
        self,
        specs: List[SeriesConfig],
        start: date,
        end: Optional[date] = None,
    ) -> List[FetchResult]:
        """Default: serial. Override for batched providers (Bloomberg)."""
        out = []
        for s in specs:
            try:
                out.append(self.fetch_one(s, start, end))
            except Exception as e:  # noqa: BLE001
                out.append(FetchResult(name=s.name, series=pd.Series(dtype=float), ok=False, error=str(e)))
        return out

    def health_check(self) -> bool:
        """Quick liveness check used by smoke_test."""
        return True
