"""One-click refresh: pulls Bloomberg + Haver deltas for both countries,
updates the Parquet cache, and regenerates Chile.xlsx and Argentina.xlsx.

Usage:
    python -m scripts.refresh_all                # delta refresh
    python -m scripts.refresh_all --force-full   # ignore cache, full re-pull
    python -m scripts.refresh_all --country CL   # one country only
"""

from __future__ import annotations

import argparse
import sys
from datetime import date, datetime
from pathlib import Path
from typing import List

# allow `python -m scripts.refresh_all` from repo root
REPO = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(REPO))

from rich.console import Console
from rich.table import Table

from src.cache.store import CacheStore
from src.config_loader import CountryConfig, SeriesConfig, Settings, load_all
from src.fetchers import get_fetcher
from src.reports.excel import build_country_workbook

console = Console()


# ────────────────────────────────────────────────────────────────────────── #
def _start_date_for(spec: SeriesConfig, settings: Settings, force_full: bool,
                    store: CacheStore) -> date:
    history = settings.history
    default_start = datetime.fromisoformat(history["default_start"]).date()
    markets_start = datetime.fromisoformat(history["markets_start"]).date()
    credit_start = datetime.fromisoformat(history["credit_start"]).date()

    if spec.tab == "credit":
        start = credit_start
    elif spec.tab == "markets":
        start = markets_start
    else:
        start = default_start

    if force_full:
        return start

    plan_start, _ = store.plan_pull_window(
        spec, default_start=start,
        always_repull_days=settings.refresh.get("always_repull_days", 30)
    )
    return plan_start


def _refresh_country(cfg: CountryConfig, settings: Settings, force_full: bool):
    console.rule(f"[bold cyan]{cfg.country}[/bold cyan]")
    cache = CacheStore(root=Path(REPO) / settings.paths["cache"], iso=cfg.iso)

    # group by provider
    by_provider: dict[str, List[SeriesConfig]] = {}
    for s in cfg.series:
        if s.provider == "derived":
            continue
        by_provider.setdefault(s.provider, []).append(s)

    all_status: dict[str, str] = {}
    for provider, specs in by_provider.items():
        console.print(f"[bold]{provider}[/bold] — {len(specs)} series")
        fetcher = get_fetcher(provider)

        # determine the earliest start across this batch (single call)
        starts = [_start_date_for(s, settings, force_full, cache) for s in specs]
        batch_start = min(starts)
        end = date.today()

        results = fetcher.fetch_many(specs, start=batch_start, end=end)
        statuses = cache.apply_results(results)
        all_status.update(statuses)

    # status table
    tbl = Table(title=f"{cfg.country} — refresh summary", show_lines=False)
    tbl.add_column("Series", overflow="fold")
    tbl.add_column("Status")
    for s in cfg.series:
        st = all_status.get(s.name, "skipped")
        style = "green" if st.startswith("ok") else ("yellow" if st == "skipped" else "red")
        tbl.add_row(s.name, f"[{style}]{st}[/]")
    console.print(tbl)

    # build the workbook
    cache_frame = cache.load_all([s.name for s in cfg.series])
    out_path = Path(REPO) / settings.paths["outputs"] / f"{cfg.country}.xlsx"
    build_country_workbook(cfg, cache_frame, settings, out_path)
    console.print(f"[bold green]✓[/] wrote {out_path}")


# ────────────────────────────────────────────────────────────────────────── #
def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--force-full", action="store_true",
                        help="ignore cache, re-pull full history")
    parser.add_argument("--country", choices=["CL", "AR"], default=None,
                        help="refresh only one country (default: both)")
    args = parser.parse_args()

    settings, countries = load_all(REPO / "config")
    targets = [countries[args.country]] if args.country else list(countries.values())

    for cfg in targets:
        _refresh_country(cfg, settings, args.force_full)

    console.rule("[bold green]done[/]")


if __name__ == "__main__":
    main()
