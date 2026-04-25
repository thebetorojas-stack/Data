"""Frequency transforms (level / yoy / mom / qoq / saar) and resampling."""

from __future__ import annotations

import numpy as np
import pandas as pd


FREQ_PANDAS = {"D": "D", "M": "ME", "Q": "QE", "A": "YE"}


def resample_to(s: pd.Series, freq: str, how: str = "last") -> pd.Series:
    """Resample to month/quarter/year-end. `how` = 'last' | 'mean' | 'sum'."""
    if s is None or s.empty:
        return s
    rule = FREQ_PANDAS.get(freq.upper(), freq)
    grouper = s.resample(rule)
    if how == "mean":
        return grouper.mean()
    if how == "sum":
        return grouper.sum(min_count=1)
    return grouper.last()


def apply_transform(s: pd.Series, transform: str, freq: str) -> pd.Series:
    """Apply a transform code to a level series.

    transform: level | yoy | mom | qoq | saar
    freq: the *target* frequency (M | Q | A | D)
    """
    if s is None or s.empty:
        return s
    t = transform.lower()
    f = freq.upper()

    if t == "level":
        return s

    if t == "yoy":
        if f == "M":
            return s.pct_change(12) * 100
        if f == "Q":
            return s.pct_change(4) * 100
        if f == "A":
            return s.pct_change(1) * 100
        # daily/other: 252 trading days as approximation
        return s.pct_change(252) * 100

    if t == "mom":
        if f != "M":
            s = resample_to(s, "M")
        return s.pct_change(1) * 100

    if t == "qoq":
        if f != "Q":
            s = resample_to(s, "Q")
        return s.pct_change(1) * 100

    if t == "saar":
        # quarterly compound growth annualized
        if f != "Q":
            s = resample_to(s, "Q")
        return ((1 + s.pct_change(1)) ** 4 - 1) * 100

    raise ValueError(f"Unknown transform: {transform}")
