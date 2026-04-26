"""em_morning_brief.py — EM macro & credit morning brief, free data, no Bloomberg.

Pulls global stress drivers from FRED (US HY OAS, EM sovereign OAS, VIX, DXY,
Treasuries, oil, S&P) and EM FX / equity / commodities from Yahoo Finance.
Computes Δ1d / Δ1w / Δ1m, z-score (vs trailing 60d daily-change SD), and 1-year
percentile so you instantly see which moves are noise vs. headlines.

Output:
  * colored table in the terminal
  * markdown brief at briefs/brief_YYYY-MM-DD.md
  * 3 PNG charts (HY OAS history, LatAm FX indexed, US 2s10s)
  * --json flag dumps the same data as JSON for scripting

Run:   python em_morning_brief.py
First-time setup:  pip install pandas requests matplotlib yfinance rich

Works on Python 3.8+. No Bloomberg, no Haver, no API keys required.
"""

from __future__ import annotations

import argparse
import json
import sys
from datetime import datetime
from io import StringIO
from pathlib import Path
from typing import Dict, List, Optional, Tuple

import pandas as pd
import requests


# ── Data sources ────────────────────────────────────────────────────────── #

FRED_BASE = "https://fred.stlouisfed.org/graph/fredgraph.csv"

# label                               (FRED id,            units,  multiplier)
FRED_SERIES: Dict[str, Tuple[str, str, float]] = {
    "US HY OAS (bps)":                ("BAMLH0A0HYM2",     "bps",  100),
    "BBB IG OAS (bps)":               ("BAMLC0A4CBBB",     "bps",  100),
    "EM HG sov OAS (bps)":            ("BAMLEMHGEMHGOAS",  "bps",  100),
    "VIX":                            ("VIXCLS",           "idx",  1),
    "USD broad (DTWEXBGS)":           ("DTWEXBGS",         "idx",  1),
    "UST 10Y (%)":                    ("DGS10",            "%",    1),
    "UST 2Y (%)":                     ("DGS2",             "%",    1),
    "WTI ($)":                        ("DCOILWTICO",       "USD",  1),
    "S&P 500":                        ("SP500",            "idx",  1),
}

YF_SERIES: Dict[str, str] = {
    "USD/ARS":  "ARS=X",
    "USD/CLP":  "CLP=X",
    "USD/BRL":  "BRL=X",
    "USD/MXN":  "MXN=X",
    "Copper":   "HG=F",
    "Soybean":  "ZS=F",
    "Merval":   "^MERV",
    "Bovespa":  "^BVSP",
}


# ── Fetchers ────────────────────────────────────────────────────────────── #

def fetch_fred(series_id: str) -> Optional[pd.Series]:
    try:
        r = requests.get(FRED_BASE, params={"id": series_id}, timeout=20)
        r.raise_for_status()
        df = pd.read_csv(StringIO(r.text))
        if df.shape[1] < 2:
            return None
        df.columns = ["date", "value"]
        df["date"] = pd.to_datetime(df["date"])
        df["value"] = pd.to_numeric(df["value"].replace(".", pd.NA), errors="coerce")
        df = df.dropna().set_index("date")
        if df.empty:
            return None
        return df["value"].rename(series_id)
    except Exception as e:
        print(f"  ! FRED {series_id}: {e}", file=sys.stderr)
        return None


def fetch_yf(ticker: str) -> Optional[pd.Series]:
    try:
        import yfinance as yf
    except ImportError:
        return None
    try:
        h = yf.Ticker(ticker).history(period="2y", auto_adjust=False)
        if h is None or h.empty or "Close" not in h.columns:
            return None
        s = h["Close"].dropna()
        # strip tz info to keep types consistent
        if s.index.tz is not None:
            s.index = s.index.tz_localize(None)
        return s.rename(ticker)
    except Exception as e:
        print(f"  ! YF {ticker}: {e}", file=sys.stderr)
        return None


# ── Stats ───────────────────────────────────────────────────────────────── #

def stat_row(label: str, s: Optional[pd.Series], units: str = "") -> Dict:
    if s is None or s.empty:
        return {"name": label, "value": None, "d1": None, "d5": None, "d22": None,
                "z": None, "pct1y": None, "units": units, "asof": None}
    s = s.sort_index().dropna()
    last = float(s.iloc[-1])

    def chg(n: int) -> Optional[float]:
        return float(last - s.iloc[-1 - n]) if len(s) > n else None

    d1, d5, d22 = chg(1), chg(5), chg(22)
    daily_changes = s.diff().tail(60).dropna()
    sd_60 = float(daily_changes.std()) if len(daily_changes) > 5 else None
    z = (d1 / sd_60) if (sd_60 and d1 is not None and sd_60 != 0) else None

    last_year = s.tail(252)
    pct1y = (float((last_year < last).sum() / max(len(last_year), 1) * 100)
             if len(last_year) else None)

    return {"name": label, "value": last, "d1": d1, "d5": d5, "d22": d22,
            "z": z, "pct1y": pct1y, "units": units,
            "asof": s.index.max().strftime("%Y-%m-%d")}


# ── Formatters ──────────────────────────────────────────────────────────── #

def fmt_value(v: Optional[float], units: str = "") -> str:
    if v is None:
        return "—"
    if units == "bps":
        return f"{v:,.0f}"
    if units == "%":
        return f"{v:.2f}%"
    if abs(v) >= 1000:
        return f"{v:,.0f}"
    return f"{v:,.2f}"


def fmt_delta(v: Optional[float]) -> str:
    if v is None:
        return "—"
    sign = "+" if v >= 0 else ""
    if abs(v) >= 100:
        return f"{sign}{v:,.0f}"
    return f"{sign}{v:,.2f}"


def fmt_z(z: Optional[float]) -> str:
    if z is None:
        return "—"
    return f"{z:+.2f}σ"


# ── Output ──────────────────────────────────────────────────────────────── #

def render_terminal(rows: List[Dict]) -> None:
    try:
        from rich.console import Console
        from rich.table import Table
    except ImportError:
        for r in rows:
            print(f"{r['name']:<24} {fmt_value(r['value'], r['units']):>10}  "
                  f"d1 {fmt_delta(r['d1']):>8}  z {fmt_z(r['z']):>8}")
        return

    console = Console()
    today = datetime.now().strftime("%Y-%m-%d")
    console.print(f"\n[bold cyan]EM Morning Brief — {today}[/]")
    console.print("[dim]Sources: FRED (free) · Yahoo Finance (free). No Bloomberg.[/]\n")

    n_fred = len(FRED_SERIES)
    sections = [
        ("Global stress drivers", rows[:n_fred]),
        ("EM FX, equity, commodities", rows[n_fred:]),
    ]
    for title, items in sections:
        tbl = Table(title=title, show_header=True, header_style="bold cyan",
                    title_justify="left", title_style="bold")
        tbl.add_column("Indicator", justify="left")
        for col in ("Current", "Δ1d", "Δ1w", "Δ1m", "z(1d)", "1y %ile"):
            tbl.add_column(col, justify="right")
        for r in items:
            row_style = ""
            if r["z"] is not None and abs(r["z"]) >= 1.5:
                row_style = "bold red" if r["z"] > 0 else "bold green"
            pct = f"{r['pct1y']:.0f}%" if r["pct1y"] is not None else "—"
            tbl.add_row(
                r["name"],
                fmt_value(r["value"], r["units"]),
                fmt_delta(r["d1"]),
                fmt_delta(r["d5"]),
                fmt_delta(r["d22"]),
                fmt_z(r["z"]),
                pct,
                style=row_style,
            )
        console.print(tbl)

    movers = [r for r in rows if r["z"] is not None and abs(r["z"]) >= 1.5]
    if movers:
        console.print("\n[bold]Headline movers (|z| ≥ 1.5):[/]")
        for r in movers:
            arrow = "▲" if r["z"] > 0 else "▼"
            console.print(f"  {arrow} {r['name']}: {fmt_delta(r['d1'])} "
                          f"({fmt_z(r['z'])}, {r['pct1y']:.0f}th %ile)")
    else:
        console.print("\n[dim]No headline movers — all moves within ±1.5σ today.[/]")


def make_charts(out_dir: Path, date_str: str) -> List[Path]:
    try:
        import matplotlib
        matplotlib.use("Agg")
        import matplotlib.pyplot as plt
    except ImportError:
        return []

    paths: List[Path] = []

    # Chart 1: HY OAS 5-year history
    hy = fetch_fred("BAMLH0A0HYM2")
    if hy is not None:
        hy = hy * 100
        fig, ax = plt.subplots(figsize=(9, 4))
        hy.tail(252 * 5).plot(ax=ax, color="#1F4E79", lw=1.4)
        # add 1-stdev bands
        rolling = hy.tail(252 * 5).rolling(60).mean()
        rolling.plot(ax=ax, color="#888", lw=0.8, ls="--", label="60d MA")
        ax.set_title("US HY OAS (bps) — 5-year history", fontsize=11, loc="left")
        ax.set_xlabel(""); ax.grid(alpha=0.25)
        ax.spines["top"].set_visible(False); ax.spines["right"].set_visible(False)
        ax.legend(frameon=False, loc="upper right", fontsize=9)
        p = out_dir / f"brief_{date_str}_hy_oas.png"
        fig.tight_layout(); fig.savefig(p, dpi=140); plt.close(fig)
        paths.append(p)

    # Chart 2: LatAm FX indexed (100 = 1y ago)
    fig, ax = plt.subplots(figsize=(9, 4))
    series_to_plot = [
        ("USD/ARS", "ARS=X", "#C8102E"),
        ("USD/CLP", "CLP=X", "#1F4E79"),
        ("USD/BRL", "BRL=X", "#548235"),
        ("USD/MXN", "MXN=X", "#7030A0"),
    ]
    plotted = 0
    for label, ticker, color in series_to_plot:
        s = fetch_yf(ticker)
        if s is None or len(s) < 252:
            continue
        s_norm = s / s.iloc[-252] * 100
        s_norm.tail(252).plot(ax=ax, lw=1.3, label=label, color=color)
        plotted += 1
    if plotted:
        ax.set_title("LatAm FX — indexed (100 = 1y ago, ↑ = local depreciation)",
                     fontsize=11, loc="left")
        ax.axhline(100, color="#000", lw=0.5, alpha=0.4)
        ax.set_xlabel(""); ax.grid(alpha=0.25)
        ax.spines["top"].set_visible(False); ax.spines["right"].set_visible(False)
        ax.legend(frameon=False, loc="upper left", fontsize=9, ncol=2)
        p = out_dir / f"brief_{date_str}_emfx.png"
        fig.tight_layout(); fig.savefig(p, dpi=140); plt.close(fig)
        paths.append(p)
    else:
        plt.close(fig)

    # Chart 3: US 2s10s slope (recession proxy)
    t10 = fetch_fred("DGS10"); t2 = fetch_fred("DGS2")
    if t10 is not None and t2 is not None:
        df = pd.concat([t10.rename("t10"), t2.rename("t2")], axis=1).dropna()
        slope = (df["t10"] - df["t2"]) * 100  # to bps
        fig, ax = plt.subplots(figsize=(9, 4))
        slope.tail(252 * 5).plot(ax=ax, color="#C00000", lw=1.4)
        ax.axhline(0, color="#000", lw=0.5, alpha=0.5)
        ax.set_title("US 2s10s slope (bps) — 5-year. Negative = inverted curve.",
                     fontsize=11, loc="left")
        ax.set_xlabel(""); ax.grid(alpha=0.25)
        ax.spines["top"].set_visible(False); ax.spines["right"].set_visible(False)
        p = out_dir / f"brief_{date_str}_2s10s.png"
        fig.tight_layout(); fig.savefig(p, dpi=140); plt.close(fig)
        paths.append(p)

    return paths


def write_markdown(rows: List[Dict], charts: List[Path],
                   out_dir: Path, date_str: str) -> Path:
    md = out_dir / f"brief_{date_str}.md"
    n_fred = len(FRED_SERIES)
    L: List[str] = [f"# EM Morning Brief — {date_str}", ""]
    L.append("_Sources: FRED · Yahoo Finance. No Bloomberg required._")
    L.append("")

    def section(title: str, items: List[Dict]) -> None:
        L.append(f"## {title}")
        L.append("")
        L.append("| Indicator | Current | Δ1d | Δ1w | Δ1m | z(1d) | 1y %ile |")
        L.append("|---|---:|---:|---:|---:|---:|---:|")
        for r in items:
            pct = f"{r['pct1y']:.0f}%" if r["pct1y"] is not None else "—"
            L.append(
                f"| {r['name']} | {fmt_value(r['value'], r['units'])} | "
                f"{fmt_delta(r['d1'])} | {fmt_delta(r['d5'])} | {fmt_delta(r['d22'])} | "
                f"{fmt_z(r['z'])} | {pct} |"
            )
        L.append("")

    section("Global stress drivers", rows[:n_fred])
    section("EM FX, equity, commodities", rows[n_fred:])

    movers = [r for r in rows if r["z"] is not None and abs(r["z"]) >= 1.5]
    if movers:
        L.append("## Headline movers (|z| ≥ 1.5)")
        L.append("")
        for r in movers:
            arrow = "▲" if r["z"] > 0 else "▼"
            L.append(f"- **{r['name']}**: {arrow} {fmt_delta(r['d1'])} "
                     f"(z = {fmt_z(r['z'])}, {r['pct1y']:.0f}th %ile)")
        L.append("")
    else:
        L.append("> No headline movers — all moves within ±1.5σ today.")
        L.append("")

    if charts:
        L.append("## Charts")
        L.append("")
        for p in charts:
            L.append(f"![{p.stem}]({p.name})")
            L.append("")

    md.write_text("\n".join(L), encoding="utf-8")
    return md


# ── Main ────────────────────────────────────────────────────────────────── #

def collect_data() -> List[Dict]:
    rows: List[Dict] = []
    print("Fetching FRED series...", file=sys.stderr)
    for label, (sid, units, mult) in FRED_SERIES.items():
        s = fetch_fred(sid)
        if s is not None:
            s = s * mult
        rows.append(stat_row(label, s, units))
    print("Fetching Yahoo Finance series...", file=sys.stderr)
    for label, ticker in YF_SERIES.items():
        s = fetch_yf(ticker)
        rows.append(stat_row(label, s, ""))
    return rows


def main() -> None:
    ap = argparse.ArgumentParser(description=__doc__.split("\n\n")[0])
    ap.add_argument("--no-charts", action="store_true",
                    help="skip PNG generation (faster)")
    ap.add_argument("--json", action="store_true",
                    help="dump rows as JSON to stdout instead of markdown/charts")
    ap.add_argument("--out-dir", default="briefs",
                    help="folder to write markdown and PNGs to (default: briefs/)")
    args = ap.parse_args()

    out_dir = Path(args.out_dir)
    out_dir.mkdir(exist_ok=True)
    date_str = datetime.now().strftime("%Y-%m-%d")

    rows = collect_data()

    if args.json:
        json.dump(rows, sys.stdout, default=str, indent=2)
        return

    render_terminal(rows)

    charts = [] if args.no_charts else make_charts(out_dir, date_str)
    md_path = write_markdown(rows, charts, out_dir, date_str)
    print(f"\nwrote {md_path}", file=sys.stderr)
    if charts:
        print(f"wrote {len(charts)} chart(s) to {out_dir}/", file=sys.stderr)


if __name__ == "__main__":
    main()
