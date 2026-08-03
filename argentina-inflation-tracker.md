# Argentina inflation tracker — build kit

Owner: Alberto (UBS). Purpose: monthly-refreshed Excel pack of INDEC CPI seasonality,
components, momentum and a bull/base/bear forecast.

## Data source
INDEC IPC nacional (base dic-2016 = 100), via the official series API. The sandbox has no
direct network egress, so **fetch with WebFetch**, not curl:

- headline: `https://apis.datos.gob.ar/series/api/series/?ids=148.3_INIVELNAL_DICI_M_26&format=csv&start_date=2018-12-01&limit=1000`
- core:     `.../?ids=148.3_INUCLEONAL_DICI_M_19&...`
- regulated:`.../?ids=148.3_IREGULANAL_DICI_M_22&...` (starts 2022-01)
- seasonal: `.../?ids=148.3_IESTACINAL_DICI_M_25&...` (starts 2022-01)

Prompt WebFetch with: "Output the raw CSV verbatim inside a code block, every row, no commentary."
Split long pulls into two calls (2018-12→2022-12, 2023-01→present) to avoid truncation.

## Validation gate (do this every rebuild)
Compounded m/m must reproduce INDEC's published Dec/Dec: 2019 53.8 · 2020 36.1 · 2021 50.9 ·
2022 94.8 · 2023 211.4 · 2024 117.8 · 2025 31.5. Jun-2026: y/y 33.5%, YTD 16.8%.
If a check fails, the transcription is wrong — re-fetch, do not ship.

## Rebuild procedure
1. Fetch the four series (above).
2. Recreate `raw.py` with the four CSV blocks as the strings `head`, `core`, `reg`, `est`
   (format `YYYY-MM,value` per line).
3. Save `build_workbook.py` (below) next to it and run it. It auto-detects the last month.
4. `python3 /root/.claude/skills/xlsx/scripts/recalc.py Argentina_inflation_seasonality.xlsx 90`
   must return `"status": "success"` with 0 errors.
5. Spot-check the Seasonality Headline tab against the published print, then SendUserFile.
6. Refresh the Forecast tab's realised months; revisit the scenario paths only if the print
   materially changes the picture, and re-check the BCRA REM (published ~6th of each month)
   for the consensus benchmark cell.

## Current call (as of Aug-2026, data through Jun-2026)
Dec/Dec 2026: bull 28.1 · base 30.8 · bear 35.8 (REM 30.0)
Dec/Dec 2027: bull 12.9 · base 18.7 · bear 28.7 (REM 11.8)
Thesis: the FX band's CPI indexation removes the nominal anchor that drove 2024-25
disinflation, so inflation converges to the crawl rather than below it. Consensus is my
bull case, not my base case.

---

## raw.py
```python
head = """2018-12,184.2552
2019-01,189.6101
2019-02,196.7501
2019-03,205.9571
2019-04,212.9596414
2019-05,219.5691
2019-06,225.537
2019-07,230.494
2019-08,239.6077
2019-09,253.7102
2019-10,262.0661
2019-11,273.2158
2019-12,283.4442
2020-01,289.8299
2020-02,295.666
2020-03,305.5515
2020-04,310.1243
2020-05,314.9087
2020-06,321.9738
2020-07,328.2014
2020-08,337.0632
2020-09,346.6207
2020-10,359.657
2020-11,371.0211
2020-12,385.8826
2021-01,401.5071
2021-02,415.8595
2021-03,435.8657
2021-04,453.6503
2021-05,468.725
2021-06,483.6049
2021-07,498.0987
2021-08,510.3942
2021-09,528.4968
2021-10,547.0802
2021-11,560.9184
2021-12,582.4575
2022-01,605.0317
2022-02,633.4341
2022-03,676.0566
2022-04,716.9399
2022-05,753.147
2022-06,793.0278
2022-07,851.761
2022-08,911.1316
2022-09,967.3076
2022-10,1028.706
2022-11,1079.2787
2022-12,1134.5875
2023-01,1202.979
2023-02,1282.7091
2023-03,1381.1601
2023-04,1497.2147
2023-05,1613.5895
2023-06,1709.6115
2023-07,1818.0838
2023-08,2044.2832
2023-09,2304.9242
2023-10,2496.273
2023-11,2816.0628
2023-12,3533.1922
2024-01,4261.5324
2024-02,4825.7881
2024-03,5357.0929
2024-04,5830.2271
2024-05,6073.7
2024-06,6351.7145
2024-07,6607.7479
2024-08,6883.4412
2024-09,7122.2421
2024-10,7313.9542
2024-11,7491.4314
2024-12,7694.0075
2025-01,7864.1257
2025-02,8052.9927
2025-03,8353.3158
2025-04,8585.6078
2025-05,8714.4871
2025-06,8855.5681
2025-07,9023.973
2025-08,9193.2441
2025-09,9384.0922
2025-10,9603.8623
2025-11,9841.3581
2025-12,10121.3715
2026-01,10413.0309
2026-02,10714.6255
2026-03,11077.0608
2026-04,11363.0904
2026-05,11607.3937
2026-06,11826.4103"""

core = """2018-12,178.9326
2019-01,184.3056
2019-02,191.4115
2019-03,200.1771
2019-04,207.7179
2019-05,214.3182
2019-06,220.1725
2019-07,224.9011
2019-08,235.3317
2019-09,250.4818
2019-10,260.0118
2019-11,270.3554
2019-12,280.3107
2020-01,287.1529
2020-02,293.9922
2020-03,303.1665
2020-04,308.3681
2020-05,313.3195
2020-06,320.5733
2020-07,328.6943
2020-08,338.6645
2020-09,346.511
2020-10,358.4989
2020-11,372.6154
2020-12,390.7655
2021-01,406.0674
2021-02,422.5932
2021-03,441.72
2021-04,461.8578
2021-05,477.9293
2021-06,495.1809
2021-07,510.6974
2021-08,526.6934
2021-09,543.8748
2021-10,561.2843
2021-11,579.5414
2021-12,605.1759
2022-01,625.1594
2022-02,652.9838
2022-03,694.7515
2022-04,741.3622
2022-05,779.9267
2022-06,819.5734
2022-07,879.6383
2022-08,939.4677
2022-09,991.2585
2022-10,1046.265
2022-11,1096.024
2022-12,1153.5771
2023-01,1215.3182
2023-02,1308.7788
2023-03,1403.1434
2023-04,1521.0163
2023-05,1640.2912
2023-06,1746.7824
2023-07,1860.2801
2023-08,2116.751
2023-09,2400.9928
2023-10,2612.5451
2023-11,2962.7206
2023-12,3799.9487
2024-01,4568.1649
2024-02,5129.5988
2024-03,5612.3991
2024-04,5965.8926
2024-05,6188.7
2024-06,6414.6575
2024-07,6657.49
2024-08,6931.7976
2024-09,7157.9271
2024-10,7366.7823
2024-11,7566.7782
2024-12,7808.7094
2025-01,7995.938
2025-02,8229.4476
2025-03,8492.1966
2025-04,8765.1966
2025-05,8957.7912
2025-06,9111.304
2025-07,9246.0737
2025-08,9435.1467
2025-09,9614.0254
2025-10,9827.1413
2025-11,10086.4793
2025-12,10392.4922
2026-01,10665.8531
2026-02,10993.8681
2026-03,11344.4566
2026-04,11603.8447
2026-05,11824.7213
2026-06,12010.3199"""

reg = """2022-01,512.3766
2022-02,528.2683
2022-03,572.7992
2022-04,595.2223
2022-05,629.0029
2022-06,662.0938
2022-07,694.6115
2022-08,738.3105
2022-09,771.8891
2022-10,829.1011
2022-11,880.7676
2022-12,925.7853
2023-01,991.5192
2023-02,1042.4951
2023-03,1128.8908
2023-04,1184.1465
2023-05,1291.1045
2023-06,1383.8634
2023-07,1476.8552
2023-08,1599.0767
2023-09,1732.073
2023-10,1846.7566
2023-11,2032.4708
2023-12,2452.7844
2024-01,3104.5367
2024-02,3759.784
2024-03,4441.3132
2024-04,5258.2523
2024-05,5469.5
2024-06,5910.154
2024-07,6166.198
2024-08,6529.3696
2024-09,6825.2577
2024-10,7009.2808
2024-11,7257.6652
2024-12,7502.5373
2025-01,7698.607
2025-02,7877.3099
2025-03,8129.994
2025-04,8273.8986
2025-05,8382.4957
2025-06,8569.0645
2025-07,8766.629
2025-08,9000.5641
2025-09,9237.6711
2025-10,9474.6115
2025-11,9748.309
2025-12,10069.0986
2026-01,10313.8016
2026-02,10752.23
2026-03,11298.7045
2026-04,11828.0421
2026-05,12111.4533
2026-06,12389.4344"""

est = """2022-01,658.2543
2022-02,713.7013
2022-03,757.8963
2022-04,799.0308
2022-05,825.8088
2022-06,880.3186
2022-07,979.7957
2022-08,1064.731
2022-09,1189.6609
2022-10,1297.2889
2022-11,1350.5167
2022-12,1413.2059
2023-01,1525.1043
2023-02,1575.9625
2023-03,1722.9662
2023-04,1940.9169
2023-05,2056.8651
2023-06,2094.7124
2023-07,2201.586
2023-08,2436.0941
2023-09,2794.0242
2023-10,3006.9009
2023-11,3392.0317
2023-12,3942.5343
2024-01,4581.5309
2024-02,4979.8267
2024-03,5534.7104
2024-04,6085.2438
2024-05,6521.2
2024-06,6811.1236
2024-07,7156.7976
2024-08,7266.119
2024-09,7479.6823
2024-10,7581.5911
2024-11,7490.4606
2024-12,7385.2224
2025-01,7411.5104
2025-02,7350.8325
2025-03,7971.2748
2025-04,8121.3082
2025-05,7899.4762
2025-06,7882.5607
2025-07,8201.9212
2025-08,8137.0132
2025-09,8316.258
2025-10,8548.5791
2025-11,8584.9379
2025-12,8637.1055
2026-01,9130.0503
2026-02,9013.6124
2026-03,9107.887
2026-04,9103.5096
2026-05,9420.5342
2026-06,9738.5609"""
```

## build_workbook.py
```python
"""
Build: Argentina INDEC inflation seasonality workbook.
Data source: INDEC IPC nacional via apis.datos.gob.ar series API.
Re-run monthly after the INDEC print (~13th-15th) to refresh.
"""
import raw
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.chart import LineChart, BarChart, Reference, Series
from openpyxl.chart.marker import Marker
from openpyxl.formatting.rule import ColorScaleRule
from openpyxl.comments import Comment

# ---------------------------------------------------------------- data
def parse(txt):
    d = {}
    for line in txt.strip().split("\n"):
        k, v = line.split(",")
        d[k] = float(v)
    return d

HEAD, CORE, REG, EST = parse(raw.head), parse(raw.core), parse(raw.reg), parse(raw.est)
MONTHS = sorted(HEAD)                       # 2018-12 .. last print
LAST = MONTHS[-1]
LAST_Y, LAST_M = int(LAST[:4]), int(LAST[5:])
YEARS = list(range(2019, LAST_Y + 1))
MLAB = ["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"]

# full month grid through Dec-2027 so new prints can simply be pasted in
GRID = []
y, m = 2018, 12
while (y, m) <= (2027, 12):
    GRID.append(f"{y}-{m:02d}")
    m += 1
    if m == 13:
        y, m = y + 1, 1

# ---------------------------------------------------------------- style
ARIAL   = "Arial"
INK     = "1A1A1A"
MUTED   = "6B6B6B"
RULE    = "D8D8D8"
INPUTC  = "0000FF"          # hardcoded inputs
ACCENT  = "1F3B63"
HDR_FILL = PatternFill("solid", fgColor="F2F4F7")
TITLE_FILL = PatternFill("solid", fgColor="1F3B63")
YEL = PatternFill("solid", fgColor="FFF6CC")

YEAR_COLORS = {2019:"D6D6D6", 2020:"BFC9D4", 2021:"A3B3C6", 2022:"879DB8",
               2023:"6B87A9", 2024:"4F719B", 2025:"2F5C8F", 2026:"D1495B",
               2027:"E8A33D"}

def title(ws, cell, text, sub=None):
    ws[cell] = text
    ws[cell].font = Font(ARIAL, 14, bold=True, color="FFFFFF")
    ws[cell].fill = TITLE_FILL
    ws[cell].alignment = Alignment(vertical="center")
    ws.row_dimensions[ws[cell].row].height = 26
    if sub:
        r = ws[cell].row + 1
        c = ws[cell].column_letter
        ws[f"{c}{r}"] = sub
        ws[f"{c}{r}"].font = Font(ARIAL, 9, italic=True, color=MUTED)

def hdr(ws, row, col, text, width=None, wrap=False):
    c = ws.cell(row=row, column=col, value=text)
    c.font = Font(ARIAL, 9, bold=True, color=INK)
    c.fill = HDR_FILL
    c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=wrap)
    c.border = Border(bottom=Side("thin", color=ACCENT))
    if width:
        ws.column_dimensions[get_column_letter(col)].width = width
    return c

def sheet_base(ws):
    ws.sheet_view.showGridLines = False
    ws.sheet_properties.tabColor = ACCENT

def line_chart(w=26, h=13):
    ch = LineChart()
    ch.width, ch.height = w, h
    ch.style = None
    ch.dispBlanksAs = "gap"
    ch.y_axis.majorGridlines.spPr = None
    return ch

def style_series(s, color, width=20000, dashed=False):
    s.smooth = False
    s.marker = Marker(symbol="none")
    s.graphicalProperties.line.solidFill = color
    s.graphicalProperties.line.width = width
    if dashed:
        s.graphicalProperties.line.dashStyle = "dash"

wb = Workbook()

# ================================================================ README
ws = wb.active
ws.title = "Read me"
sheet_base(ws)
ws.column_dimensions["A"].width = 2
ws.column_dimensions["B"].width = 26
ws.column_dimensions["C"].width = 105
title(ws, "B2", "Argentina inflation — INDEC IPC seasonality pack")
ws["B3"] = f"Data through {LAST}. Built {'automatically' if True else ''} from INDEC national CPI."
ws["B3"].font = Font(ARIAL, 9, italic=True, color=MUTED)

readme = [
    ("What's in here", ""),
    ("Data", "Every INDEC national CPI series used here, as index levels (blue = hardcoded from source) with month-on-month, year-on-year, 3m annualised and year-to-date calculated by formula."),
    ("Seasonality Headline", "The chart you asked for: x-axis Jan-Dec, one line per year 2019 onwards, values = monthly headline m/m."),
    ("Seasonality Core", "Same, for IPC Núcleo (core)."),
    ("Heatmap", "Year x month grid, colour-scaled — the fastest way to see which months run hot in a given regime."),
    ("Momentum", "3-month annualised vs year-on-year, headline and core. Shows turning points ~9 months before the y/y does."),
    ("Components", "Núcleo vs Regulados vs Estacionales — where the 2026 pressure actually sits."),
    ("Seasonal Factors", "Each month's m/m divided by that year's average month. Scale-free, so 2023 (211%) and 2025 (31%) are comparable. 1.20x = that month runs 20% hotter than a typical month of the same year."),
    ("Forecast", "Bull / base / bear monthly paths to Dec-2027 with the FX and tariff assumptions behind each, benchmarked to the BCRA REM."),
    ("", ""),
    ("Source", "INDEC, Índice de precios al consumidor nacional (base dic-2016 = 100), retrieved from the official series API at apis.datos.gob.ar (series 148.3_INIVELNAL_DICI_M_26, 148.3_INUCLEONAL_DICI_M_19, 148.3_IREGULANAL_DICI_M_22, 148.3_IESTACINAL_DICI_M_25)."),
    ("Validation", "Compounded monthly rates reproduce INDEC's published Dec/Dec prints exactly: 2019 53.8%, 2020 36.1%, 2021 50.9%, 2022 94.8%, 2023 211.4%, 2024 117.8%, 2025 31.5%. June-2026 y/y 33.5% and YTD 16.8% also match the official release."),
    ("", ""),
    ("How to update", ""),
    ("Option 1 — automatic", "A scheduled task re-runs this build every month on the 16th (INDEC publishes around the 13th-15th) and sends you the refreshed workbook. Nothing to do."),
    ("Option 2 — by hand", "On the Data tab, find the first empty yellow row and paste the new index levels into columns E-H. Every m/m, y/y, matrix and heatmap cell updates by formula. Chart series ranges cover the full calendar year, so a part-year line will read zero for months you have not filled — the automatic rebuild handles that for you."),
    ("Caution", "INDEC occasionally revises the seasonal and regulated sub-indices. The automatic rebuild re-pulls the whole history each month, so revisions flow through; a manual paste-in does not restate history."),
]
r = 5
for k, v in readme:
    if v == "":
        ws.cell(row=r, column=2, value=k).font = Font(ARIAL, 11, bold=True, color=ACCENT)
    else:
        ws.cell(row=r, column=2, value=k).font = Font(ARIAL, 10, bold=True, color=INK)
        c = ws.cell(row=r, column=3, value=v)
        c.font = Font(ARIAL, 10, color=INK)
        c.alignment = Alignment(wrap_text=True, vertical="top")
        ws.row_dimensions[r].height = 15 * (1 + len(v) // 105)
    r += 1

# ================================================================ DATA
ws = wb.create_sheet("Data")
sheet_base(ws)
title(ws, "B2", "INDEC IPC nacional — index levels and derived rates",
      "Blue = hardcoded index level from INDEC. Black = formula. Paste new prints into the yellow rows.")

HR = 5              # header row
R0 = 6              # first data row
cols = [("Month", 10), ("Year", 7), ("M#", 5),
        ("Nivel General\n(index)", 13), ("Núcleo\n(index)", 12),
        ("Regulados\n(index)", 12), ("Estacionales\n(index)", 13),
        ("m/m\nNivel Gen.", 11), ("m/m\nNúcleo", 10), ("m/m\nRegulados", 11),
        ("m/m\nEstacion.", 11), ("y/y\nNivel Gen.", 11), ("y/y\nNúcleo", 10),
        ("3m ann.\nNivel Gen.", 11), ("3m ann.\nNúcleo", 11), ("YTD\nNivel Gen.", 11)]
for i, (t, w) in enumerate(cols):
    hdr(ws, HR, 2 + i, t, w, wrap=True)
ws.row_dimensions[HR].height = 30
ws.column_dimensions["A"].width = 2

SRC = {"E": HEAD, "F": CORE, "G": REG, "H": EST}
for i, ym in enumerate(GRID):
    r = R0 + i
    yy, mm = int(ym[:4]), int(ym[5:])
    ws.cell(row=r, column=2, value=ym).font = Font(ARIAL, 9, color=INK)
    ws.cell(row=r, column=3, value=yy).font = Font(ARIAL, 9, color=MUTED)
    ws.cell(row=r, column=3).number_format = "0"
    ws.cell(row=r, column=4, value=mm).font = Font(ARIAL, 9, color=MUTED)
    for col_letter, src in SRC.items():
        col = ord(col_letter) - 64
        c = ws.cell(row=r, column=col)
        if ym in src:
            c.value = src[ym]
            c.font = Font(ARIAL, 9, color=INPUTC)
        else:
            c.font = Font(ARIAL, 9, color=INPUTC)
            c.fill = YEL
        c.number_format = "#,##0.00"
    # derived
    p = r - 1
    for j, col_letter in enumerate("EFGH"):
        c = ws.cell(row=r, column=9 + j)
        if r > R0:
            c.value = f'=IF(OR({col_letter}{r}="",{col_letter}{p}=""),"",{col_letter}{r}/{col_letter}{p}-1)'
        c.number_format = "0.00%"
        c.font = Font(ARIAL, 9, color=INK)
    for j, col_letter in enumerate("EF"):          # y/y
        c = ws.cell(row=r, column=13 + j)
        if r >= R0 + 12:
            c.value = f'=IF(OR({col_letter}{r}="",{col_letter}{r-12}=""),"",{col_letter}{r}/{col_letter}{r-12}-1)'
        c.number_format = "0.0%"
        c.font = Font(ARIAL, 9, color=INK)
    for j, col_letter in enumerate("EF"):          # 3m annualised
        c = ws.cell(row=r, column=15 + j)
        if r >= R0 + 3:
            c.value = f'=IF(OR({col_letter}{r}="",{col_letter}{r-3}=""),"",({col_letter}{r}/{col_letter}{r-3})^4-1)'
        c.number_format = "0.0%"
        c.font = Font(ARIAL, 9, color=INK)
    c = ws.cell(row=r, column=17)                  # YTD
    if yy >= 2019:
        c.value = (f'=IF(E{r}="","",IFERROR(E{r}/INDEX($E:$E,MATCH(C{r}-1&"-12",$B:$B,0))-1,""))')
    c.number_format = "0.0%"
    c.font = Font(ARIAL, 9, color=INK)

LASTROW = R0 + len(GRID) - 1
ws.freeze_panes = f"B{R0}"
ws["B2"].comment = Comment(
    "Index levels are INDEC published values (base dic-2016 = 100) pulled from apis.datos.gob.ar. "
    "Everything to the right of column G is calculated in-sheet.", "Claude")

DATA = "Data"

# ================================================================ seasonality sheets
def seasonality_sheet(name, mm_col_letter, idx_col_letter, label, note):
    ws = wb.create_sheet(name)
    sheet_base(ws)
    title(ws, "B2", f"Seasonality — {label}: month-on-month % by calendar year", note)
    ws.column_dimensions["A"].width = 2
    ws.column_dimensions["B"].width = 11
    hdr(ws, 5, 2, "Month", 11)
    for j, yy in enumerate(YEARS):
        hdr(ws, 5, 3 + j, yy, 10).number_format = "0"
    for i, ml in enumerate(MLAB):
        r = 6 + i
        c = ws.cell(row=r, column=2, value=ml)
        c.font = Font(ARIAL, 9, bold=True, color=INK)
        for j, yy in enumerate(YEARS):
            col = 3 + j
            L = get_column_letter(col)
            cc = ws.cell(row=r, column=col)
            cc.value = (
                f'=IF(COUNTIFS(\'{DATA}\'!$C${R0}:$C${LASTROW},{L}$5,'
                f'\'{DATA}\'!$D${R0}:$D${LASTROW},{i+1},'
                f'\'{DATA}\'!${idx_col_letter}${R0}:${idx_col_letter}${LASTROW},">0")=0,"",'
                f'SUMIFS(\'{DATA}\'!${mm_col_letter}${R0}:${mm_col_letter}${LASTROW},'
                f'\'{DATA}\'!$C${R0}:$C${LASTROW},{L}$5,'
                f'\'{DATA}\'!$D${R0}:$D${LASTROW},{i+1}))')
            cc.number_format = "0.0%"
            cc.font = Font(ARIAL, 9, color=INK)
    # year summary rows
    ws.cell(row=19, column=2, value="Dec/Dec*").font = Font(ARIAL, 9, bold=True, color=ACCENT)
    ws.cell(row=20, column=2, value="Avg month").font = Font(ARIAL, 9, bold=True, color=ACCENT)
    for j, yy in enumerate(YEARS):
        L = get_column_letter(3 + j)
        endm = f"{yy}-12" if yy < LAST_Y else LAST
        c1 = ws.cell(row=19, column=3 + j, value=(
            f'=IFERROR(INDEX(\'{DATA}\'!${idx_col_letter}:${idx_col_letter},MATCH("{endm}",\'{DATA}\'!$B:$B,0))/'
            f'INDEX(\'{DATA}\'!${idx_col_letter}:${idx_col_letter},MATCH("{yy-1}-12",\'{DATA}\'!$B:$B,0))-1,"")'))
        c1.number_format = "0.0%"
        c1.font = Font(ARIAL, 9, bold=True, color=ACCENT)
        c2 = ws.cell(row=20, column=3 + j, value=f'=IFERROR(AVERAGE({L}6:{L}17),"")')
        c2.number_format = "0.00%"
        c2.font = Font(ARIAL, 9, color=MUTED)
    ws.cell(row=21, column=2,
            value=f"* {LAST_Y} column is year-to-date through {LAST}, not a full year.")
    ws.cell(row=21, column=2).font = Font(ARIAL, 8, italic=True, color=MUTED)

    ch = line_chart()
    ch.title = f"{label} — monthly inflation by year"
    ch.y_axis.title = "% month-on-month"
    ch.y_axis.numFmt = "0.0%"
    cats = Reference(ws, min_col=2, min_row=6, max_row=17)
    for j, yy in enumerate(YEARS):
        maxr = 17
        if yy == LAST_Y:
            maxr = 5 + LAST_M
        ref = Reference(ws, min_col=3 + j, min_row=5, max_row=maxr)
        s = Series(ref, title_from_data=True)
        style_series(s, YEAR_COLORS[yy],
                     width=32000 if yy >= LAST_Y - 1 else 17000)
        ch.series.append(s)
    ch.set_categories(cats)
    ws.add_chart(ch, "B24")

    ch2 = line_chart()
    ch2.title = f"{label} — same lines, y-axis capped at 8% (2023-24 spikes clipped)"
    ch2.y_axis.numFmt = "0.0%"
    ch2.y_axis.scaling.min = 0
    ch2.y_axis.scaling.max = 0.08
    for j, yy in enumerate(YEARS):
        maxr = 17 if yy != LAST_Y else 5 + LAST_M
        ref = Reference(ws, min_col=3 + j, min_row=5, max_row=maxr)
        s2 = Series(ref, title_from_data=True)
        style_series(s2, YEAR_COLORS[yy], width=32000 if yy >= LAST_Y - 1 else 17000)
        ch2.series.append(s2)
    ch2.set_categories(cats)
    ws.add_chart(ch2, "B50")
    return ws

seasonality_sheet("Seasonality Headline", "I", "E", "IPC Nivel General (headline)",
                  "One line per calendar year, Jan-Dec on the x-axis. Recent years are drawn darker; 2026 in red.")
seasonality_sheet("Seasonality Core", "J", "F", "IPC Núcleo (core)",
                  "Core strips out regulated tariffs and seasonal items — the cleanest read on underlying momentum.")

# ================================================================ HEATMAP
ws = wb.create_sheet("Heatmap")
sheet_base(ws)
title(ws, "B2", "Monthly inflation heatmap — year x month",
      "Green = cool, red = hot, scaled within each block. Headline on top, core below.")
ws.column_dimensions["A"].width = 2
ws.column_dimensions["B"].width = 11

def heat_block(start_row, src_sheet, label):
    ws.cell(row=start_row, column=2, value=label).font = Font(ARIAL, 11, bold=True, color=ACCENT)
    hdr(ws, start_row + 1, 2, "Year", 11)
    for j, ml in enumerate(MLAB):
        hdr(ws, start_row + 1, 3 + j, ml, 7)
    hdr(ws, start_row + 1, 15, "Dec/Dec", 10)
    for i, yy in enumerate(YEARS):
        r = start_row + 2 + i
        c = ws.cell(row=r, column=2, value=yy)
        c.font = Font(ARIAL, 9, bold=True, color=INK)
        c.number_format = "0"
        for j in range(12):
            cc = ws.cell(row=r, column=3 + j,
                         value=f"='{src_sheet}'!{get_column_letter(3+i)}{6+j}")
            cc.number_format = "0.0%"
            cc.font = Font(ARIAL, 9, color=INK)
            cc.alignment = Alignment(horizontal="center")
        cc = ws.cell(row=r, column=15, value=f"='{src_sheet}'!{get_column_letter(3+i)}19")
        cc.number_format = "0.0%"
        cc.font = Font(ARIAL, 9, bold=True, color=ACCENT)
    rng = f"C{start_row+2}:N{start_row+1+len(YEARS)}"
    ws.conditional_formatting.add(rng, ColorScaleRule(
        start_type="min", start_color="63BE7B",
        mid_type="percentile", mid_value=50, mid_color="FFEB84",
        end_type="max", end_color="F8696B"))
    return start_row + 3 + len(YEARS)

nxt = heat_block(5, "Seasonality Headline", "Headline (IPC Nivel General)")
heat_block(nxt + 1, "Seasonality Core", "Core (IPC Núcleo)")
ws.cell(row=nxt + len(YEARS) + 6, column=2,
        value="Colour scale is relative to the whole block, so 2023-24 dominates the red end. "
              "For a like-for-like read across regimes use the Seasonal Factors tab.").font = Font(ARIAL, 8, italic=True, color=MUTED)

# ================================================================ MOMENTUM
ws = wb.create_sheet("Momentum")
sheet_base(ws)
title(ws, "B2", "Momentum — 3-month annualised vs year-on-year",
      "3m annualised turns first. Where it sits below y/y, the y/y is still falling mechanically.")
row_2019 = R0 + GRID.index("2019-01")
row_last = R0 + GRID.index(LAST)

ch = line_chart(30, 14)
ch.title = "Headline and core: 3m annualised vs y/y"
ch.y_axis.numFmt = "0%"
ch.y_axis.title = "%"
cats = Reference(wb[DATA], min_col=2, min_row=row_2019, max_row=row_last)
specs = [(15, "3m ann. headline", "D1495B", 30000, False),
         (16, "3m ann. core", "E8A33D", 26000, False),
         (13, "y/y headline", "2F5C8F", 22000, True),
         (14, "y/y core", "8AA0B8", 20000, True)]
for col, nm, colr, wdt, dash in specs:
    ref = Reference(wb[DATA], min_col=col, min_row=row_2019 - 1, max_row=row_last)
    s = Series(ref, title=nm)
    style_series(s, colr, wdt, dash)
    ch.series.append(s)
ch.set_categories(cats)
ws.add_chart(ch, "B5")

ch2 = line_chart(30, 12)
ch2.title = "Zoom: since Jan-2025"
ch2.y_axis.numFmt = "0%"
row_25 = R0 + GRID.index("2025-01")
cats2 = Reference(wb[DATA], min_col=2, min_row=row_25, max_row=row_last)
for col, nm, colr, wdt, dash in specs:
    ref = Reference(wb[DATA], min_col=col, min_row=row_25, max_row=row_last)
    s = Series(ref, title=nm)
    style_series(s, colr, wdt, dash)
    ch2.series.append(s)
ch2.set_categories(cats2)
ws.add_chart(ch2, "B34")

# ================================================================ COMPONENTS
ws = wb.create_sheet("Components")
sheet_base(ws)
title(ws, "B2", "Where the pressure sits — núcleo vs regulados vs estacionales",
      "Regulated prices (tariffs, transport, fuel, health plans) are the 2026 story.")
row_22 = R0 + GRID.index("2022-01")
ch = line_chart(30, 13)
ch.title = "Monthly % change by component"
ch.y_axis.numFmt = "0%"
cats = Reference(wb[DATA], min_col=2, min_row=row_22, max_row=row_last)
for col, nm, colr, wdt in [(9, "Nivel General", "1A1A1A", 30000),
                           (10, "Núcleo", "2F5C8F", 24000),
                           (11, "Regulados", "D1495B", 24000),
                           (12, "Estacionales", "9AB3C9", 18000)]:
    ref = Reference(wb[DATA], min_col=col, min_row=row_22 - 1, max_row=row_last)
    s = Series(ref, title=nm)
    style_series(s, colr, wdt)
    ch.series.append(s)
ch.set_categories(cats)
ws.add_chart(ch, "B5")

# cumulative-change blocks
comp_years = [y for y in YEARS if y >= 2023]
comp_rows = [("Nivel General", "E"), ("Núcleo", "F"), ("Regulados", "G"), ("Estacionales", "H")]

def comp_block(tr, heading, note, end_of, chart_title, anchor, partial_flag):
    ws.cell(row=tr, column=2, value=heading).font = Font(ARIAL, 11, bold=True, color=ACCENT)
    hdr(ws, tr + 1, 2, "Component", 16)
    for j, yy in enumerate(comp_years):
        lbl = ("%d YTD" % yy) if (partial_flag and yy == LAST_Y) else yy
        h = hdr(ws, tr + 1, 3 + j, lbl, 11)
        if not isinstance(lbl, str):
            h.number_format = "0"
    for i, (nm, cl) in enumerate(comp_rows):
        r = tr + 2 + i
        ws.cell(row=r, column=2, value=nm).font = Font(ARIAL, 9, bold=True, color=INK)
        for j, yy in enumerate(comp_years):
            c = ws.cell(row=r, column=3 + j, value=(
                f'=IFERROR(INDEX(\'{DATA}\'!${cl}:${cl},MATCH("{end_of(yy)}",\'{DATA}\'!$B:$B,0))/'
                f'INDEX(\'{DATA}\'!${cl}:${cl},MATCH("{yy-1}-12",\'{DATA}\'!$B:$B,0))-1,"")'))
            c.number_format = "0.0%"
            c.font = Font(ARIAL, 9, color=INK)
    n = ws.cell(row=tr + 6, column=2, value=note)
    n.font = Font(ARIAL, 8, italic=True, color=MUTED)
    n.alignment = Alignment(wrap_text=False, vertical="top")
    bc = BarChart()
    bc.type, bc.width, bc.height = "col", 22, 11
    bc.title = chart_title
    bc.y_axis.numFmt = "0%"
    bc.dispBlanksAs = "gap"
    bc.add_data(Reference(ws, min_col=3, min_row=tr + 1, max_col=2 + len(comp_years), max_row=tr + 5),
                titles_from_data=True, from_rows=False)
    bc.set_categories(Reference(ws, min_col=2, min_row=tr + 2, max_row=tr + 5))
    ws.add_chart(bc, anchor)

comp_block(
    32,
    "1. Full calendar year — cumulative from December of the prior year",
    f"Each bar compounds that year's monthly rates off a Dec = 100 base. The {LAST_Y} bar covers only "
    f"Jan-{MLAB[LAST_M-1]} ({LAST_M} of 12 months), so it is NOT comparable with the full years beside it. Use block 2 for that.",
    lambda yy: (f"{yy}-12" if yy < LAST_Y else LAST),
    "Cumulative change by component — full year (current year is part-year)",
    "H32", True)

comp_block(
    41,
    f"2. Like-for-like — January to {MLAB[LAST_M-1]} of each year",
    f"The same {LAST_M}-month window in every year, so these bars ARE directly comparable. "
    f"Read this one to judge whether {LAST_Y} is running hotter or cooler than {LAST_Y-1}.",
    lambda yy: f"{yy}-{LAST_M:02d}",
    f"Cumulative change by component — Jan to {MLAB[LAST_M-1]}, like-for-like",
    "H58", False)

# ================================================================ SEASONAL FACTORS
ws = wb.create_sheet("Seasonal Factors")
sheet_base(ws)
title(ws, "B2", "Seasonal factors — each month vs its own year's average month",
      "Scale-free ratio, so 2023 (211% inflation) and 2025 (31%) are directly comparable. 1.20x = 20% hotter than a typical month that year.")
ws.column_dimensions["A"].width = 2
ws.column_dimensions["B"].width = 11
full_years = [y for y in YEARS if y < LAST_Y]
hdr(ws, 5, 2, "Month", 11)
for j, yy in enumerate(full_years):
    hdr(ws, 5, 3 + j, yy, 9).number_format = "0"
n = len(full_years)
hdr(ws, 5, 3 + n, "Average", 10)
hdr(ws, 5, 4 + n, "Median", 10)
hdr(ws, 5, 5 + n, "ex-2020", 10)
for i, ml in enumerate(MLAB):
    r = 6 + i
    ws.cell(row=r, column=2, value=ml).font = Font(ARIAL, 9, bold=True, color=INK)
    for j, yy in enumerate(full_years):
        L = get_column_letter(3 + j)
        c = ws.cell(row=r, column=3 + j, value=(
            f'=IFERROR(\'Seasonality Headline\'!{L}{r}/\'Seasonality Headline\'!{L}$20,"")'))
        c.number_format = "0.00x"
        c.font = Font(ARIAL, 9, color=INK)
    ce = get_column_letter(2 + n)
    ws.cell(row=r, column=3 + n, value=f"=IFERROR(AVERAGE(C{r}:{ce}{r}),\"\")").number_format = "0.00x"
    ws.cell(row=r, column=4 + n, value=f"=IFERROR(MEDIAN(C{r}:{ce}{r}),\"\")").number_format = "0.00x"
    # ex-2020 (pandemic distortions)
    if 2020 in full_years:
        k = full_years.index(2020)
        Lk = get_column_letter(3 + k)
        expr = f'=IFERROR((SUM(C{r}:{ce}{r})-{Lk}{r})/({n}-1),"")'
    else:
        expr = f'=IFERROR(AVERAGE(C{r}:{ce}{r}),"")'
    ws.cell(row=r, column=5 + n, value=expr).number_format = "0.00x"
    for col in (3 + n, 4 + n, 5 + n):
        ws.cell(row=r, column=col).font = Font(ARIAL, 9, bold=True, color=ACCENT)

bc = BarChart()
bc.type, bc.width, bc.height = "col", 26, 11
bc.title = "Average seasonal factor by month (2019 – %d, headline)" % (LAST_Y - 1)
bc.y_axis.numFmt = "0.00"
bc.dispBlanksAs = "gap"
bc.add_data(Reference(ws, min_col=3 + n, min_row=5, max_row=17), titles_from_data=True)
bc.set_categories(Reference(ws, min_col=2, min_row=6, max_row=17))
bc.series[0].graphicalProperties.solidFill = "2F5C8F"
ws.add_chart(bc, "B20")
ws.cell(row=19, column=2,
        value="Read: a factor above 1.00x means the month is systematically hotter than the year's average month. "
              "Denominator is that year's own average monthly rate (Seasonality tab, row 20).").font = Font(ARIAL, 8, italic=True, color=MUTED)

# ================================================================ FORECAST
BASE26 = [1.9, 1.8, 1.8, 1.9, 1.9, 2.1]
BULL26 = [1.7, 1.6, 1.5, 1.5, 1.4, 1.6]
BEAR26 = [2.2, 2.4, 2.5, 2.7, 2.6, 2.8]
BASE27 = [2.0, 1.8, 1.7, 1.6, 1.4, 1.3, 1.3, 1.2, 1.2, 1.2, 1.2, 1.4]
BULL27 = [1.5, 1.3, 1.3, 1.1, 1.0, 0.9, 0.9, 0.8, 0.8, 0.8, 0.8, 1.0]
BEAR27 = [2.8, 2.6, 2.5, 2.3, 2.1, 2.0, 1.9, 1.9, 1.8, 1.8, 1.8, 2.0]

ws = wb.create_sheet("Forecast")
sheet_base(ws)
title(ws, "B2", "Forecast — Argentina CPI, bull / base / bear to Dec-2027",
      "Monthly paths are judgement inputs (blue). Dec/Dec figures are formulas. Benchmarked to the BCRA REM of June-2026.")
ws.column_dimensions["A"].width = 2
ws.column_dimensions["B"].width = 11

hdr(ws, 5, 2, "Month", 11)
for j, nm in enumerate(["Actual", "Bull", "Base", "Bear"]):
    hdr(ws, 5, 3 + j, nm, 11)
for j, nm in enumerate(["Bull index", "Base index", "Bear index"]):
    hdr(ws, 5, 7 + j, nm, 12, wrap=True)

fc_months = [g for g in GRID if g >= "2025-01"]
scen = {}
M26 = ["2026-%02d" % m for m in range(7, 13)]
M27 = ["2027-%02d" % m for m in range(1, 13)]
FCAST = {}
for _ms, _b, _ba, _be in ((M26, BULL26, BASE26, BEAR26), (M27, BULL27, BASE27, BEAR27)):
    for _m, _x, _y2, _z in zip(_ms, _b, _ba, _be):
        FCAST[_m] = (_x, _y2, _z)
for i, ym in enumerate(fc_months):
    scen[ym] = (None, None, None) if ym <= LAST else FCAST.get(ym, (None, None, None))

FR0 = 6
for i, ym in enumerate(fc_months):
    r = FR0 + i
    ws.cell(row=r, column=2, value=ym).font = Font(ARIAL, 9, color=INK)
    a = ws.cell(row=r, column=3)
    a.value = f'=IFERROR(IF(INDEX(\'{DATA}\'!$I:$I,MATCH($B{r},\'{DATA}\'!$B:$B,0))="","",INDEX(\'{DATA}\'!$I:$I,MATCH($B{r},\'{DATA}\'!$B:$B,0))),"")'
    a.number_format = "0.0%"
    a.font = Font(ARIAL, 9, color="008000")
    for j, v in enumerate(scen[ym]):
        c = ws.cell(row=r, column=4 + j)
        if ym == LAST:
            c.value = f"=$C{r}"          # anchor so forecast lines join the actual
            c.font = Font(ARIAL, 9, color="008000")
        elif v is not None:
            c.value = v / 100
            c.font = Font(ARIAL, 9, color=INPUTC)
        c.number_format = "0.0%"
    for j, L in enumerate("DEF"):
        c = ws.cell(row=r, column=7 + j)
        if ym <= LAST:
            c.value = f'=IFERROR(INDEX(\'{DATA}\'!$E:$E,MATCH($B{r},\'{DATA}\'!$B:$B,0)),"")'
        else:
            c.value = f'={get_column_letter(7+j)}{r-1}*(1+{L}{r})'
        c.number_format = "#,##0"
        c.font = Font(ARIAL, 8, color=MUTED)
FRL = FR0 + len(fc_months) - 1

# results block
res = FRL + 2
ws.cell(row=res, column=2, value="Dec/Dec outcome").font = Font(ARIAL, 11, bold=True, color=ACCENT)
labels = [("2026 (Dec/Dec)", 2026), ("2027 (Dec/Dec)", 2027)]
hdr(ws, res + 1, 2, "", 11)
for j, nm in enumerate(["Bull", "Base", "Bear"]):
    hdr(ws, res + 1, 4 + j, nm, 11)
hdr(ws, res + 1, 7, "BCRA REM (Jun-26 survey)", 24)

r26_start = FR0 + fc_months.index("2026-01")
r26_end = FR0 + fc_months.index("2026-12")
r27_start = FR0 + fc_months.index("2027-01")
r27_end = FR0 + fc_months.index("2027-12")
for i, (nm, yy) in enumerate(labels):
    r = res + 2 + i
    ws.cell(row=r, column=2, value=nm).font = Font(ARIAL, 10, bold=True, color=INK)
    s, e = (r26_start, r26_end) if yy == 2026 else (r27_start, r27_end)
    for j in range(3):
        P = get_column_letter(7 + j)
        if yy == LAST_Y:
            c = ws.cell(row=r, column=4 + j, value=(
                f'={P}${e}/INDEX(\'{DATA}\'!$E:$E,MATCH("{LAST_Y-1}-12",\'{DATA}\'!$B:$B,0))-1'))
        else:
            c = ws.cell(row=r, column=4 + j, value=f'={P}${e}/{P}${e-12}-1')
        c.number_format = "0.0%"
        c.font = Font(ARIAL, 11, bold=True, color=ACCENT)
ws.cell(row=res + 2, column=7, value=0.30).number_format = "0.0%"
ws.cell(row=res + 3, column=7, value=0.118).number_format = "0.0%"
for r in (res + 2, res + 3):
    ws.cell(row=r, column=7).font = Font(ARIAL, 10, color=INPUTC)
ws.cell(row=res + 4, column=2,
        value="REM = BCRA Relevamiento de Expectativas de Mercado, June-2026 round (44 participants, published 6-Jul-2026): "
              "30% for 2026; the 2027 median is the last published reading. 2026 formula splices realised months with the scenario path.").font = Font(ARIAL, 8, italic=True, color=MUTED)

# assumptions
asr = res + 7
ws.cell(row=asr, column=2, value="Scenario assumptions").font = Font(ARIAL, 11, bold=True, color=ACCENT)
hdr(ws, asr + 1, 2, "Driver", 22)
for j, nm in enumerate(["Bull", "Base", "Bear"]):
    hdr(ws, asr + 1, 3 + j, nm, 34, wrap=True)
assump = [
    ("FX — official, Dec-26",
     "ARS 1,600 — spot drifts up less than the band crawl; real appreciation resumes",
     "ARS 1,690 — spot tracks the band's CPI-indexed crawl, roughly 2.2%/m from ARS 1,470",
     "ARS 1,950 — band ceiling tested; a step devaluation of 12-15% in Q4"),
    ("FX pass-through",
     "0.10 within a quarter — import competition holds, margins absorb",
     "0.15 within two quarters — the 2024-25 realised range",
     "0.30 — a discrete jump reprices tradables fast"),
    ("Regulated tariffs",
     "Catch-up essentially complete; regulated converges toward core by Q4-26",
     "Regulated keeps running ~50-80bp/m above core through mid-27, then converges",
     "Energy and transport subsidy removal re-accelerates; regulated stays 150bp/m above core"),
    ("Wages / indexation",
     "Paritarias settle at or below expected inflation; backward-looking indexation breaks",
     "Paritarias track trailing CPI — inertia keeps a floor near 1.2-1.4%/m",
     "Real wage recovery demands re-open; formal wages index to trailing 12m"),
    ("Fiscal / monetary",
     "Primary surplus above 1.5% GDP, no monetary financing, remonetisation orderly",
     "Primary surplus around 1.0-1.5% GDP maintained; peso demand grows with activity",
     "Surplus slips ahead of the 2027 electoral cycle; peso demand stalls"),
    ("Key risk to this view",
     "Disinflation faster than any Argentine episode since convertibility — historically rare",
     "The band's CPI indexation removes the nominal anchor that drove 2024-25; inflation converges to the crawl, not below it",
     "A band break is the single event that resets the whole path — watch reserve accumulation, not the CPI"),
]
for i, row in enumerate(assump):
    r = asr + 2 + i
    ws.cell(row=r, column=2, value=row[0]).font = Font(ARIAL, 9, bold=True, color=INK)
    ws.cell(row=r, column=2).alignment = Alignment(vertical="top", wrap_text=True)
    for j in range(3):
        c = ws.cell(row=r, column=3 + j, value=row[1 + j])
        c.font = Font(ARIAL, 9, color=INK)
        c.alignment = Alignment(wrap_text=True, vertical="top")
    ws.row_dimensions[r].height = 44

ch = line_chart(30, 14)
ch.title = "Monthly CPI: realised and scenario paths"
ch.y_axis.numFmt = "0.0%"
cats = Reference(ws, min_col=2, min_row=FR0, max_row=FRL)
for col, nm, colr, wdt, dash in [(3, "Actual", "1A1A1A", 32000, False),
                                 (4, "Bull", "4F9D69", 22000, True),
                                 (5, "Base", "2F5C8F", 26000, True),
                                 (6, "Bear", "D1495B", 22000, True)]:
    ref = Reference(ws, min_col=col, min_row=FR0 - 1, max_row=FRL)
    s = Series(ref, title=nm)
    style_series(s, colr, wdt, dash)
    ch.series.append(s)
ch.set_categories(cats)
ws.add_chart(ch, "K5")

wb.save("Argentina_inflation_seasonality.xlsx")
print("saved")
```
