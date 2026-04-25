"""Load and validate YAML config files."""

from __future__ import annotations

from dataclasses import dataclass, field
from pathlib import Path
from typing import Any, Dict, List, Optional

import yaml


@dataclass
class SeriesConfig:
    name: str
    label: str
    provider: str          # bloomberg | haver | derived
    ticker: str
    frequency: str         # D | M | Q | A
    tab: str               # monthly | quarterly | annual | credit | markets
    category: str
    units: str
    transform: str = "level"
    chart: bool = False
    field: str = "PX_LAST"
    curve_maturity_years: Optional[float] = None

    @classmethod
    def from_dict(cls, d: Dict[str, Any]) -> "SeriesConfig":
        return cls(
            name=d["name"],
            label=d["label"],
            provider=d["provider"].lower(),
            ticker=d.get("ticker", "") or "",
            frequency=d.get("frequency", "M").upper(),
            tab=d["tab"],
            category=d.get("category", ""),
            units=d.get("units", "level"),
            transform=d.get("transform", "level"),
            chart=bool(d.get("chart", False)),
            field=d.get("field", "PX_LAST"),
            curve_maturity_years=d.get("curve_maturity_years"),
        )


@dataclass
class CountryConfig:
    country: str
    iso: str
    currency: str
    series: List[SeriesConfig] = field(default_factory=list)

    @classmethod
    def load(cls, path: str | Path) -> "CountryConfig":
        with open(path, "r", encoding="utf-8") as f:
            raw = yaml.safe_load(f)
        return cls(
            country=raw["country"],
            iso=raw["iso"],
            currency=raw["currency"],
            series=[SeriesConfig.from_dict(s) for s in raw["series"]],
        )

    def by_tab(self, tab: str) -> List[SeriesConfig]:
        return [s for s in self.series if s.tab == tab]

    def by_name(self, name: str) -> Optional[SeriesConfig]:
        for s in self.series:
            if s.name == name:
                return s
        return None


@dataclass
class Settings:
    paths: Dict[str, str]
    history: Dict[str, str]
    refresh: Dict[str, Any]
    excel: Dict[str, Any]
    dashboard: Dict[str, Any]

    @classmethod
    def load(cls, path: str | Path) -> "Settings":
        with open(path, "r", encoding="utf-8") as f:
            raw = yaml.safe_load(f)
        return cls(**raw)


def load_all(config_dir: str | Path = "config"):
    """Convenience: returns (settings, {iso: CountryConfig})."""
    config_dir = Path(config_dir)
    settings = Settings.load(config_dir / "settings.yaml")
    countries = {}
    for fname in ("chile.yaml", "argentina.yaml"):
        cc = CountryConfig.load(config_dir / fname)
        countries[cc.iso] = cc
    return settings, countries
