# Macro & Hard-Currency Credit — Chile / Argentina

Two self-updating Excel workbooks driven by Bloomberg formulas. **No Python required.**

## What's in this folder

- `Chile.xlsx` — macro + credit + markets, all tabs, native Excel charts
- `Argentina.xlsx` — same structure for Argentina
- `EMBI_Fair_Value_Model.xlsx` — top-down regression + bottom-up country attribution model for the EMBI Global Diversified spread (3m / 6m / 12m forecasts)
- `_python_version_optional/` — the older Python+Streamlit version, kept here as a backup option only. Ignore unless you specifically want it later.

## How to use

1. Make sure Bloomberg Terminal is open and you're logged in.
2. Open `Chile.xlsx` (or `Argentina.xlsx`).
3. On the **Bloomberg** tab in the Excel ribbon, click **Refresh Workbook** (or just press F9). Data + charts update in place.
4. That's it. Email the workbook to whoever needs it — the charts retain the last refresh, so recipients without Bloomberg still see your numbers and charts.

## What's in each workbook

Seven tabs:

- **Read me** — the same instructions as above, inside the file.
- **Monthly** — CPI, activity (IMACEC/EMAE), industrial production, retail sales, unemployment, trade balance, reserves, policy rate, fiscal monthly. Charts in a 2-column grid below the data.
- **Quarterly** — real GDP YoY, GDP QoQ saar, current account % GDP, investment (GFCF), private consumption.
- **Annual** — GDP per capita USD, total public debt % GDP, external debt % GDP.
- **Credit** — EMBI country spread, CEMBI country spread, EMBI total return, three points on the USD sovereign curve (5Y / 10Y / 30Y for Chile; GD30 / GD35 / GD41 for Argentina).
- **Markets** — FX, equity index, key commodity (copper for Chile, soybeans for Argentina), 10Y local rate.
- **Tickers** — every series and its Bloomberg ticker on one filterable list. Edit a ticker here, refresh, the chart updates. This is your one place to swap things out.

## Verifying / fixing tickers

Every cell with `=BDH(...)` is a normal Excel formula. If a series shows `#N/A` or a wrong number, it's almost always one of:

- Ticker is slightly off for your subscription / region. Open Bloomberg, type `ECO <GO>` for macro indicators or `SECF <GO>` for securities, find the right ticker, paste it on the **Tickers** tab.
- Field name is wrong (`PX_LAST` vs `YLD_YTM_MID`). The Tickers tab shows which field is being used per row.
- History parameter is too aggressive. The formulas pull from 2000–2018 depending on frequency; if a series only goes back to 2018, the older years just return blank — that's normal.

## Why no Python anymore

The Python+Streamlit version had nice features (auto-cache, dashboard, programmatic refresh), but on a locked-down corporate machine with restricted Python and gated package installs, it's more friction than it's worth. The Excel + BDH approach gives you the same charts, same data, refreshes with one click, and uses tools you already have running on the machine. If you ever want the dashboard back, the full Python version is preserved in `_python_version_optional/`.
