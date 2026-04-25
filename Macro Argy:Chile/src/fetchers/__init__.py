from __future__ import annotations

from .base import Fetcher, FetchResult
from .bloomberg import BloombergFetcher
from .haver import HaverFetcher

__all__ = ["Fetcher", "FetchResult", "BloombergFetcher", "HaverFetcher"]


def get_fetcher(provider: str) -> Fetcher:
    """Factory: pick the fetcher implementation for a provider string."""
    p = provider.lower()
    if p == "bloomberg":
        return BloombergFetcher()
    if p == "haver":
        return HaverFetcher()
    raise ValueError(f"Unknown provider: {provider}")
