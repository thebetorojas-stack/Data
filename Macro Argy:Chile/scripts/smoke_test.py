"""Smoke test: verifies Bloomberg and Haver connectivity independent of the pipeline.

Usage:
    python -m scripts.smoke_test                   # tests both providers
    python -m scripts.smoke_test --provider bloomberg
    python -m scripts.smoke_test --provider haver
    python -m scripts.smoke_test --series cpi_headline_yoy --country CL
"""

from __future__ import annotations

import argparse
import sys
from datetime import date, timedelta
from pathlib import Path

REPO = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(REPO))

from rich.console import Console

from src.config_loader import load_all
from src.fetchers import get_fetcher

console = Console()


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--provider", choices=["bloomberg", "haver"], default=None)
    parser.add_argument("--series", default=None,
                        help="single series name to test (requires --country)")
    parser.add_argument("--country", choices=["CL", "AR"], default=None)
    args = parser.parse_args()

    _settings, countries = load_all(REPO / "config")

    if args.series:
        if not args.country:
            console.print("[red]--series requires --country[/]")
            sys.exit(2)
        cfg = countries[args.country]
        spec = cfg.by_name(args.series)
        if not spec:
            console.print(f"[red]No such series: {args.series}[/]")
            sys.exit(2)
        fetcher = get_fetcher(spec.provider)
        end = date.today()
        start = end - timedelta(days=120)
        r = fetcher.fetch_one(spec, start=start, end=end)
        if r.ok:
            console.print(f"[green]OK[/] {spec.name}: {len(r.series)} obs, last={r.series.iloc[-1]:.4f} on {r.series.index.max().date()}")
        else:
            console.print(f"[red]FAIL[/] {spec.name}: {r.error}")
        return

    providers = [args.provider] if args.provider else ["bloomberg", "haver"]
    for p in providers:
        f = get_fetcher(p)
        ok = f.health_check()
        marker = "[green]OK[/]" if ok else "[red]FAIL[/]"
        console.print(f"{marker} {p} health check")


if __name__ == "__main__":
    main()
