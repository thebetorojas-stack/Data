# EM Credit & Macro — Chile / Argentina

One-click refresh of macro and hard-currency credit coverage for Chile and Argentina.
Data flows from Bloomberg (`blpapi` / `xbbg`) and Haver Analytics (`Haver` Python API) into
a local Parquet cache, then into:

- **`outputs/Chile.xlsx`** and **`outputs/Argentina.xlsx`** — client-ready workbooks with native Excel charts (Monthly / Quarterly / Annual / Credit tabs).
- **Streamlit dashboard** — interactive, browser-based, for your own monitoring.

Refresh everything with a double-click on `run_refresh.bat` (Windows) or `run_refresh.sh` (Mac/Linux).

---

## 1. First-time setup (do this once)

### 1a. Install Python dependencies

```bash
pip install -r requirements.txt
```

### 1b. Bloomberg

The Bloomberg Terminal must be running and logged in on the same machine. The Desktop API
must be enabled (it is by default for most Terminal installs).

Test:

```bash
python -m scripts.smoke_test --provider bloomberg
```

If it fails: open the Terminal, run `WAPI <GO>`, confirm Desktop API is enabled, restart Python.

### 1c. Haver Analytics

Haver DLX must be installed locally. The Python API ships with DLX. If `import Haver` fails,
add the Haver Python directory to `PYTHONPATH` (typically `C:\DLX\Data` on Windows).

Test:

```bash
python -m scripts.smoke_test --provider haver
```

### 1d. Verify ticker maps

Open `config/chile.yaml` and `config/argentina.yaml`. Tickers are named with standard
conventions but **Haver database codes vary by subscription** — verify the database
prefix (e.g. `CHILE`, `ARGENT`, `EMERGE`) matches what you subscribe to.

---

## 2. Daily use

### Refresh everything

Double-click **`run_refresh.bat`** (or `./run_refresh.sh`).

This pulls only the delta since the last fetch, updates the Parquet cache, and regenerates
both Excel files. Typical run time after the first full pull: 10–30 seconds.

### Open the dashboard

Double-click **`run_dashboard.bat`**. Opens at http://localhost:8501.

The dashboard reads from the same cache, so it's instant after a refresh.

### Force a full re-pull (rare)

```bash
python -m scripts.refresh_all --force-full
```

---

## 3. Project layout

```
.
├── config/
│   ├── settings.yaml          # paths, history depth, refresh windows
│   ├── chile.yaml             # Chile ticker map (Haver + Bloomberg)
│   └── argentina.yaml         # Argentina ticker map
├── src/
│   ├── fetchers/
│   │   ├── base.py            # abstract fetcher
│   │   ├── bloomberg.py       # blpapi / xbbg wrapper
│   │   └── haver.py           # Haver Python API wrapper
│   ├── cache/
│   │   └── store.py           # Parquet cache with delta logic
│   ├── transform/
│   │   ├── frequencies.py     # M / Q / A resampling
│   │   └── credit.py          # spread / OAS / curve helpers
│   ├── reports/
│   │   ├── excel.py           # openpyxl with native charts
│   │   └── styles.py          # consistent number/color formatting
│   └── dashboard/
│       └── app.py             # Streamlit entry point
├── scripts/
│   ├── refresh_all.py         # main entry point
│   ├── refresh_country.py     # single-country refresh
│   └── smoke_test.py          # validates Bloomberg / Haver connectivity
├── cache/                     # parquet files (auto-created)
├── outputs/                   # generated .xlsx files
├── run_refresh.bat            # Windows one-click refresh
├── run_refresh.sh             # Mac/Linux one-click refresh
├── run_dashboard.bat          # Windows launch dashboard
└── run_dashboard.sh           # Mac/Linux launch dashboard
```

---

## 4. Adding or changing tickers

1. Edit `config/chile.yaml` or `config/argentina.yaml`.
2. Each entry has: `name`, `provider` (bloomberg|haver), `ticker`, `field`, `frequency` (D|M|Q|A), `category`, `chart` (true|false), `units`, `transform` (optional: `yoy`, `mom`, `level`).
3. Re-run `run_refresh.bat`. New series flow into the cache and the Excel layout picks them up automatically.

---

## 5. What's in each tab

**Monthly** — CPI (headline + core, MoM/YoY), IMACEC/EMAE, IP, retail sales, unemployment, trade balance, reserves, FX, policy rate, monetary aggregates, fiscal monthly, confidence indices.

**Quarterly** — GDP (real, by sector and expenditure), current account, BoP, FDI, gross fixed capital formation, debt/GDP.

**Annual** — GDP per capita, external debt stock and structure, sovereign debt amortization profile, fiscal medium-term framework.

**Credit** — Sovereign USD curve (yields by maturity), EMBI Global Diversified country spread (level + Δ), CEMBI Broad Diversified country, total return indices.

Each tab has a top summary block, a data table, and several charts. All charts are
**native Excel chart objects** — when the underlying data range updates, the chart updates.
You can edit colors, ranges, and titles in Excel and they'll persist if you don't change
the layout (see "Customizing the layout" in the README appendix).

---

## 6. Troubleshooting quickly

| Symptom | Likely cause | Fix |
|---|---|---|
| `Connection refused` on Bloomberg | Terminal not running | Launch Bloomberg, log in, retry |
| `ModuleNotFoundError: Haver` | Haver Python path missing | `set PYTHONPATH=C:\DLX\Data;%PYTHONPATH%` |
| One series shows `NaN` | Bad ticker or wrong Haver DB code | Edit YAML, run `smoke_test --series <name>` |
| Refresh slow on first run | Full history pull, expected | Subsequent runs use deltas (~10–30s) |
| Excel chart broken | Manual chart edit conflicted with regeneration | Delete the .xlsx, re-run refresh |

---

## 7. Conventions

- All series stored UTC date-indexed in Parquet.
- Frequency resampling: month-end (`M`), quarter-end (`Q`), year-end (`A`).
- Argentina FX: official, MEP, CCL (blue chip swap), and brecha (gap %) all tracked separately.
- Argentina fiscal: primary and overall, both BCRA and Tesoro perimeter where available.
- Spreads quoted in bps; yields in %; FX in local per USD.
