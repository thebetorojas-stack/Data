"""Streamlit dashboard.

Run with:  streamlit run src/dashboard/app.py
Or:         double-click run_dashboard.bat (Windows) / run_dashboard.sh (Mac/Linux)

Reads from the same Parquet cache as the Excel writer, so charts are instant
after a refresh. The "Refresh data" button shells out to refresh_all.py.
"""

from __future__ import annotations

import subprocess
import sys
from datetime import datetime
from pathlib import Path

import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st

# repo root resolved from this file location
REPO = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(REPO))

from src.cache.store import CacheStore  # noqa: E402
from src.config_loader import load_all  # noqa: E402
from src.transform.credit import build_curve, curve_history  # noqa: E402
from src.transform.frequencies import apply_transform, resample_to  # noqa: E402

st.set_page_config(page_title="EM Macro & Credit", layout="wide", page_icon="🌎")

settings, countries = load_all(REPO / "config")

# ─── sidebar ────────────────────────────────────────────────────────────── #
st.sidebar.title("EM Macro & Credit")
country_choice = st.sidebar.selectbox(
    "Country",
    options=list(countries.keys()),
    format_func=lambda iso: countries[iso].country,
    index=0,
)
cfg = countries[country_choice]
panel = st.sidebar.radio("Panel", ["Monthly", "Quarterly", "Annual", "Credit", "Markets"], index=0)

st.sidebar.markdown("---")
if st.sidebar.button("🔄 Refresh data"):
    with st.spinner("Pulling fresh data from Bloomberg & Haver..."):
        proc = subprocess.run(
            [sys.executable, "-m", "scripts.refresh_all"],
            cwd=str(REPO),
            capture_output=True, text=True,
        )
    if proc.returncode == 0:
        st.sidebar.success("Refresh complete")
    else:
        st.sidebar.error("Refresh failed — see console")
        with st.sidebar.expander("stderr"):
            st.code(proc.stderr[-2000:])

st.sidebar.markdown(f"_Last refresh: {datetime.now().strftime('%Y-%m-%d %H:%M')}_")

# ─── load cache ─────────────────────────────────────────────────────────── #
store = CacheStore(root=REPO / settings.paths["cache"], iso=cfg.iso)
all_names = [s.name for s in cfg.series]
cache_df = store.load_all(all_names)

if cache_df.empty:
    st.warning("Cache is empty. Click **Refresh data** in the sidebar to do an initial pull.")
    st.stop()


# ─── header ─────────────────────────────────────────────────────────────── #
st.title(f"{cfg.country} — {panel}")
st.caption(f"Currency: {cfg.currency} • Refresh window through "
           f"{cache_df.index.max().strftime('%Y-%m-%d')}")


# ─── panel router ───────────────────────────────────────────────────────── #
def panel_specs(tab: str):
    return [s for s in cfg.series if s.tab == tab.lower()]


def chart(spec, df: pd.DataFrame):
    name = spec.name
    if name not in df.columns:
        st.info(f"{spec.label} — no data yet")
        return
    s = df[name].dropna()
    if s.empty:
        st.info(f"{spec.label} — no data yet")
        return
    fig = px.line(
        s.reset_index().rename(columns={"index": "Date", name: spec.label}),
        x="Date", y=spec.label, title=spec.label,
    )
    fig.update_layout(height=320, margin=dict(l=10, r=10, t=40, b=10),
                      showlegend=False)
    fig.update_traces(line=dict(width=2))
    st.plotly_chart(fig, use_container_width=True)


def build_panel_df(tab_name: str) -> tuple[pd.DataFrame, list]:
    specs = panel_specs(tab_name)
    target_freq = {"Monthly": "M", "Quarterly": "Q", "Annual": "A",
                   "Credit": "D", "Markets": "D"}[tab_name]
    cols = {}
    for s in specs:
        if s.name not in cache_df.columns:
            continue
        ser = cache_df[s.name].dropna()
        if ser.empty:
            continue
        if target_freq != "D":
            ser = resample_to(ser, target_freq)
        ser = apply_transform(ser, s.transform, target_freq)
        cols[s.name] = ser
    if not cols:
        return pd.DataFrame(), specs
    df = pd.concat(cols.values(), axis=1, keys=cols.keys()).sort_index()
    return df, specs


# ─── render panel ───────────────────────────────────────────────────────── #
if panel == "Credit":
    # special layout: spreads + curve
    df, specs = build_panel_df(panel)

    st.subheader("Spreads")
    spread_specs = [s for s in specs if s.category == "spreads"]
    cols = st.columns(min(3, max(1, len(spread_specs))))
    for i, sp in enumerate(spread_specs):
        with cols[i % len(cols)]:
            chart(sp, df)

    st.subheader("USD Sovereign Curve")
    curve_now = build_curve(cfg, cache_df)
    if curve_now.empty:
        st.info("No curve data yet — verify USD bond tickers in the country YAML.")
    else:
        hist = curve_history(cfg, cache_df, snapshots=4)
        fig = go.Figure()
        if not hist.empty:
            for snap in hist["snapshot"].unique():
                sub = hist[hist["snapshot"] == snap]
                fig.add_trace(go.Scatter(
                    x=sub["maturity_years"], y=sub["yield_pct"],
                    mode="lines+markers", name=snap,
                ))
        else:
            fig.add_trace(go.Scatter(
                x=curve_now["maturity_years"], y=curve_now["yield_pct"],
                mode="lines+markers", name="Today",
            ))
        fig.update_layout(
            height=420, xaxis_title="Maturity (years)", yaxis_title="Yield (%)",
            margin=dict(l=10, r=10, t=30, b=10),
        )
        st.plotly_chart(fig, use_container_width=True)

    st.subheader("Total returns")
    tr_specs = [s for s in specs if s.category == "total_return"]
    for sp in tr_specs:
        chart(sp, df)

elif panel == "Markets":
    df, specs = build_panel_df(panel)
    if cfg.iso == "AR":
        # FX gap chart
        if "usdars_official" in df.columns and "usdars_ccl" in df.columns:
            from src.transform.credit import brecha as _brecha
            gap = _brecha(df["usdars_official"], df["usdars_ccl"])
            fig = go.Figure()
            fig.add_trace(go.Scatter(x=gap.index, y=gap.values, mode="lines",
                                     name="Brecha CCL/Oficial (%)"))
            fig.update_layout(title="Brecha CCL vs Oficial",
                              height=320, margin=dict(l=10, r=10, t=40, b=10))
            st.plotly_chart(fig, use_container_width=True)

    cols = st.columns(2)
    for i, sp in enumerate(s for s in specs if s.chart):
        with cols[i % 2]:
            chart(sp, df)

else:
    df, specs = build_panel_df(panel)

    # KPI strip
    kpi_specs = [s for s in specs if s.chart][:6]
    if kpi_specs:
        kpi_cols = st.columns(len(kpi_specs))
        for col, sp in zip(kpi_cols, kpi_specs):
            if sp.name in df.columns:
                v = df[sp.name].dropna()
                if not v.empty:
                    last = v.iloc[-1]
                    prev = v.iloc[-2] if len(v) > 1 else None
                    delta = (last - prev) if prev is not None and pd.notna(prev) else None
                    col.metric(sp.label, f"{last:,.2f}",
                               f"{delta:+,.2f}" if delta is not None else None)

    # chart grid
    cols = st.columns(2)
    chart_specs = [s for s in specs if s.chart]
    for i, sp in enumerate(chart_specs):
        with cols[i % 2]:
            chart(sp, df)

    # data table at bottom
    with st.expander("Show data table"):
        st.dataframe(df.tail(48), use_container_width=True)
