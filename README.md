# Macro Tracker — multi-country

A self-refreshing Excel workbook per country, covering GDP, Inflation, Fiscal, Balance of Payments, and Reserves. You edit one sheet (Codes) per country. One command (`python refresh.py`) repopulates everything — data, growth metrics, 2-year forecasts, and charts.

Currently configured: **Argentina** and **Chile**. Adding a third country is one CSV file.

## What you get

One workbook per country (`Macro_Tracker_Argentina.xlsx`, `Macro_Tracker_Chile.xlsx`), each with these tabs:

- **README** — in-workbook quick reference.
- **Codes** — the only thing you ever edit. One row per indicator. Yellow cells = inputs.
- **GDP / Inflation / Fiscal / BoP / Reserves** — one tab per category. Each indicator on the Codes sheet gets its own block on the matching tab: a quarterly Table (Date, Level, QoQ %, YoY %, Linear Forecast, Holt-Winters Forecast), an annual Table (Year, Level, YoY %, Forecast), and a chart that overlays history with both forecasts.
- **Dashboard** — one row per indicator: latest value, latest YoY %, 2-year forecast, source.

The data and forecast tables are real Excel Tables (ListObjects), and the charts reference Table columns. When `refresh.py` rewrites them, the charts pick up the new range automatically — no chart fiddling, ever.

## The minimal workflow

1. Open `Macro_Tracker_Argentina.xlsx` (or `_Chile`).
2. Edit the **Codes** sheet — change a Haver code, add/remove a row.
3. Save and close.
4. Run `python refresh.py` (rebuilds both countries) or `python refresh.py argentina` (just one).
5. Reopen the workbook. Every tab is current.

To add a third country: drop a `codes_<name>.csv` next to `refresh.py`. It'll be picked up automatically.

## Codes sheet schema

| Column | Notes |
| --- | --- |
| Section | Must be one of GDP, Inflation, Fiscal, BoP, Reserves. Drives which tab the indicator lands on. |
| Indicator | Display name (becomes the block title and chart title). |
| Country | Free text. Useful when you track multiple countries. |
| Frequency | D / W / M / Q / A — your hint about the native frequency. |
| Quarterly Code | Haver code for the quarterly view. |
| Annual Code | Haver code for the annual view. Optional — if blank the workbook aggregates from the quarterly. |
| Units | Free text, surfaces in the block header and chart Y-axis. |
| Notes | Free text. |

## Aggregation rules

Higher-frequency series get rolled up automatically:

- **GDP / Inflation** — averaged across periods (the right rule for SAAR levels and price indices).
- **Fiscal flows / BoP flows** — summed across periods.
- **Stocks (Federal Debt, Reserves)** — last value of period.

If you ever need to override one of these, change the Indicator name to include the right keyword (`debt`, `index`, `deflator`) — the override list is at the top of `macro_tracker.py`.

## Forecast methods

Both run on the quarterly series, then the annual forecast is aggregated from the quarterly:

- **Linear trend** — log-linear regression on the full history (positive series) or OLS on levels (deficit/CA series). Extends 8 quarters.
- **Holt-Winters** — additive damped-trend exponential smoothing, with seasonal terms when history allows. 95% bands are residual-σ × √h.

## Setup

```bash
pip install pandas numpy openpyxl
pip install statsmodels   # optional, upgrades the Holt-Winters fit; pure-numpy fallback otherwise
pip install haver         # Windows only — needs Haver DLX installed locally with a valid licence
```

## Data sources, in order

`refresh.py` tries each, uses whichever works:

1. **Haver DLX** via the `haver` Python package. Requires DLX installed locally (Windows + Haver subscription).
2. **CSV fallback** — drop `data/<CODE>.csv` (`date,value`) for any code. Lets Mac/Linux users iterate without DLX.
3. **Synthetic demo** — plausible-looking values per indicator family. The Source column on the Dashboard shows which path each indicator took, and the README section title in the workbook flags the demo.

## Notes for Argentina & Chile

The starter Haver codes in `codes_argentina.csv` and `codes_chile.csv` are best-guess mnemonics following Haver's IFS-style country-code conventions (Argentina = 213, Chile = 228). Every row is tagged `[VERIFY]` in Notes — please confirm each one in your Haver subscription before relying on the numbers. The structure of the workbook will work regardless; only the data fetch depends on the codes being correct.

For Argentina specifically, the linear-trend forecast is fitted on the **last 40 quarters only** (configurable in `macro_tracker.py`), so structural breaks like the 2007–2015 INDEC suspension or recent inflation regime changes don't poison the projection. Holt-Winters still uses the full history. Even so, expect inflation forecasts to be wide — that's a feature, not a bug, given the underlying volatility.

## Files in this folder

| File | Purpose |
| --- | --- |
| `Macro_Tracker_Argentina.xlsx` | Argentina workbook. Open this to see AR data. |
| `Macro_Tracker_Chile.xlsx` | Chile workbook. Open this to see CL data. |
| `codes_argentina.csv` | Starter codes for Argentina (used on first run; after that the workbook's Codes sheet is the source of truth). |
| `codes_chile.csv` | Starter codes for Chile. |
| `refresh.py` | One-command refresh. Bare = both countries; pass `argentina` / `chile` to refresh just one. |
| `macro_tracker.py` | Workbook builder (categories, blocks, Tables, charts). |
| `haver_metrics.py` | Fetcher + forecaster. Used by macro_tracker. |
| `data/` | Optional CSV cache, one file per Haver code. |
