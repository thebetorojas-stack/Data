# -*- coding: utf-8 -*-
"""
=============================================================================
 FX CARRY PAYOFF CHART  —  reusable template
=============================================================================

 What it draws
 -------------
 12-month total return on a long-high-yielder forward/NDF position, plotted
 against terminal spot, with the breakeven cushion shaded and scenario points
 marked. The point of the chart is that carry is fixed at inception, so
 terminal spot is the only variable and the whole payoff is one curve.

 How to run
 ----------
     pip install matplotlib          # the only dependency
     python fx_carry_chart.py

 Outputs three files into OUTDIR:
     <name>.svg   <- INSERT THIS INTO WORD.  True vector, infinite resolution,
                     stays sharp at any zoom and when the doc is printed or
                     distilled to PDF. Word 2016 / Microsoft 365 supports it
                     natively: Insert > Pictures > This Device > select .svg.
     <name>.png   <- 600 dpi raster fallback for older Word, PowerPoint on
                     locked-down builds, or anywhere SVG is rejected.
     <name>.pdf   <- vector, for LaTeX or for sending to design.

 If SVG is greyed out in your Word build, use the PNG: at 600 dpi and a
 physical width set by FIG_W_IN, it is already well past the point where the
 printer becomes the limiting factor. Do NOT resize the PNG upward in Word —
 change FIG_W_IN here and re-run instead, so the pixels are generated rather
 than interpolated.

 Adapting to another currency
 ----------------------------
 Everything you need is in the CONFIG block. The two things that actually
 matter are QUOTE (which way round the pair is quoted) and DIRECTION (which
 leg you are long). Get those right and the sign conventions take care of
 themselves. Worked examples for MXN, BRL, TRY and EUR are at the bottom.
=============================================================================
"""

import matplotlib
matplotlib.use("Agg")               # file output only; no interactive window
import matplotlib.pyplot as plt
from matplotlib.ticker import FuncFormatter
import numpy as np
import os

# =============================================================== CONFIG =====
# --- the trade ---------------------------------------------------------------
PAIR        = "USDCOP"      # cosmetic; used in the axis label
BASE_CCY    = "COP"         # the currency you are long / short per DIRECTION
SPOT        = 3152.0        # current spot, in the units named by QUOTE
RATE_FCY    = 12.5911       # % p.a. — local rate.  Bloomberg CLNI12M BGN Curncy
RATE_USD    = 4.02          # % p.a. — funding leg. USOSFR1 BGN Curncy
                            #   NEVER use LIBOR: ICE ceased synthetic USD
                            #   publication 30 Sep 2024. Term SOFR = TSFR3M /
                            #   TSFR1M; 1y OIS = USOSFR1.
HORIZON_Y   = 1.0           # years to the forward date. 1.0 = 12m. 0.25 = 3m.

QUOTE       = "FCY_PER_USD"
#   "FCY_PER_USD"  -> pair is quoted as units of local per 1 USD, and a HIGHER
#                     number means a WEAKER local currency.
#                     COP, MXN, BRL, CLP, TRY, ZAR, JPY, INR, IDR, PEN, PHP.
#   "USD_PER_FCY"  -> pair is quoted as USD per 1 unit of local, and a HIGHER
#                     number means a STRONGER local currency.
#                     EUR, GBP, AUD, NZD.

DIRECTION   = "LONG_FCY"    # "LONG_FCY" = long the local currency, funded in
                            # USD (the normal carry trade when RATE_FCY is the
                            # higher rate). "SHORT_FCY" flips every sign, for
                            # when you are funding IN the local currency.

FWD_METHOD  = "LINEAR"
#   "LINEAR" -> F = S x (1 + carry x h).  Matches the desk convention of simply
#               adding the quoted differential. Use this to tie to a sheet that
#               was built that way.
#   "EXACT"  -> F = S x (1+r_fcy x h)/(1+r_usd x h).  Covered interest parity.
#               Diverges from LINEAR as rates rise: at 12.6% vs 4.0% over 12m
#               the two differ by about 10 big figures on USDCOP, so pick one
#               deliberately and say which in the footnote.
#   Neither isolates the cross-currency basis. For a deliverable NDF P&L, drive
#   F off the quoted NDF points instead and set FWD_OVERRIDE below.
FWD_OVERRIDE = None         # e.g. 3422.0 to hard-set the outright and ignore
                            # FWD_METHOD entirely. None = compute it.

IMPLIED_VOL = 13.0          # % p.a. — optional. Set to None to hide the
                            # carry-to-vol annotation. Below ~0.5 the trade is
                            # thin; above ~1.0 it is genuinely attractive.

# --- scenarios ---------------------------------------------------------------
# (label, terminal spot, probability, place text ABOVE the marker?)
# Probabilities are used for the expected-return line in the subtitle. They do
# not need to sum to 1 — if they don't, the script says so rather than
# silently normalising.
SCENARIOS = [
    ("Bull", 2950.0, 0.20, True),
    ("Base", 3250.0, 0.55, False),
    ("Bear", 3950.0, 0.25, True),
]

# --- reference verticals -----------------------------------------------------
# (level, label — use \n for a second line, colour key, linestyle)
# Colour keys: "ink" "red" "green" "navy" "mut" "amber".
# Set to [] for none. The spot and breakeven lines are added automatically
# unless AUTO_REFS is False.
AUTO_REFS   = True
EXTRA_REFS  = [
    (3700.0, "Start-2026\n~3,700", "mut", ":"),
]

# --- axes --------------------------------------------------------------------
# Set any of these to None to derive it from the data.
X_MIN, X_MAX = 2800.0, 4300.0
X_STEP       = 250.0        # gap between x tick labels
Y_MIN, Y_MAX = -24.0, 26.0  # in percent
Y_STEP       = 10.0
X_DECIMALS   = 0            # 0 for COP/JPY/CLP; 4 for EURUSD; 2 for MXN

# --- text --------------------------------------------------------------------
TITLE    = "12-month total return on a long-COP NDF, by terminal USDCOP"
SUBTITLE = None             # None = auto-build from carry, breakeven, E[R]
XLABEL   = "Terminal USDCOP"
CUSHION_LABEL = None        # None = auto, e.g. "8.6% cushion"

# --- output ------------------------------------------------------------------
OUTDIR   = "."
OUTNAME  = "fx_carry_chart"
FIG_W_IN = 9.75             # inches. Match your Word text column: US Letter
                            # with 1" margins = 6.5", A4 with 2.5cm = 6.3".
                            # Drawing WIDER than the column and letting Word
                            # scale down is correct — it makes every line and
                            # glyph finer. 9.75" into a 6.5" column is a 0.67x
                            # reduction and is what the note used.
FIG_H_IN = 5.00
DPI_PNG  = 600              # 600 is print-grade. 300 is the floor. Above 600
                            # the file bloats with no visible gain.
SAVE     = ("svg", "png", "pdf")

# --- style -------------------------------------------------------------------
FONTS = ["Frutiger 45 Light", "Frutiger", "Arial", "Helvetica", "DejaVu Sans"]
#   matplotlib walks this list and takes the first one installed. Frutiger is
#   the UBS face; if it is not on the machine you silently fall back to Arial,
#   which is the correct behaviour.

C = {                       # one accent, one positive, one negative, one grey
    "red":   "#E60000",     # UBS red
    "green": "#2E7D5B",
    "navy":  "#1F3B63",
    "ink":   "#16202E",
    "mut":   "#6B7A8F",
    "amber": "#C08A2E",
    "grid":  "#E3E8EF",
}
CURVE_COLOUR   = "navy"
FILL_ALPHA     = 0.09       # fill under the curve. Keep under 0.15 — shading
BAND_ALPHA     = 0.12       # should register, not shout.
CURVE_LW       = 2.6        # one heavy line for the answer; everything thinner
SHOW_LEGS      = False      # True adds dashed spot-return and carry-return
                            # lines. Useful for a desk audience, noise for a
                            # client one.

# ============================================================== ENGINE ======
# Nothing below here needs editing for a routine currency swap, but it is all
# plain and commented so you can.

h = HORIZON_Y
carry = RATE_FCY - RATE_USD          # % p.a., local minus funding

# --- forward outright --------------------------------------------------------
# In FCY_PER_USD terms a higher-yielding local currency trades at a HIGHER
# forward (it depreciates in the forward), which is what gives you the cushion.
# In USD_PER_FCY terms it trades at a LOWER forward, so the arithmetic inverts.
if FWD_OVERRIDE is not None:
    FWD = float(FWD_OVERRIDE)
elif QUOTE == "FCY_PER_USD":
    FWD = (SPOT * (1 + carry / 100 * h) if FWD_METHOD == "LINEAR"
           else SPOT * (1 + RATE_FCY / 100 * h) / (1 + RATE_USD / 100 * h))
elif QUOTE == "USD_PER_FCY":
    FWD = (SPOT * (1 - carry / 100 * h) if FWD_METHOD == "LINEAR"
           else SPOT * (1 + RATE_USD / 100 * h) / (1 + RATE_FCY / 100 * h))
else:
    raise ValueError("QUOTE must be 'FCY_PER_USD' or 'USD_PER_FCY'")

SIGN = 1.0 if DIRECTION == "LONG_FCY" else -1.0


def total_return(S):
    """Total return in % on the funding currency's notional.

    FCY_PER_USD: you sold USD forward at FWD and buy it back at S, so the
      P&L per unit of USD notional is (FWD - S) / S. Dividing by S rather than
      FWD is what makes losses convex — this is exactly the term the naive
      additive decomposition carry + (S0/S - 1) gets wrong.
    USD_PER_FCY: you bought FCY forward at FWD and sell at S, giving
      (S - FWD) / FWD.
    """
    S = np.asarray(S, dtype=float)
    if QUOTE == "FCY_PER_USD":
        return SIGN * (FWD - S) / S * 100.0
    return SIGN * (S - FWD) / FWD * 100.0


def spot_return(S):
    """The pure currency move, holding carry aside."""
    S = np.asarray(S, dtype=float)
    if QUOTE == "FCY_PER_USD":
        return SIGN * (SPOT / S - 1) * 100.0
    return SIGN * (S / SPOT - 1) * 100.0


def carry_return(S):
    """Carry actually earned. NOT a constant.

    The coupon is fixed in local currency, but you realise it at the TERMINAL
    exchange rate, so its USD value scales by SPOT/S. If the currency halves,
    so does the carry. Charting carry as a flat line is the single commonest
    error in these decompositions and it always flatters the bear.
    """
    return total_return(S) - spot_return(S)


# Carry you actually EARN, after DIRECTION. Negative means the forward works
# against you: long EUR vs USD at these rates pays a premium rather than
# collecting one, and the band is a cost to make up, not a buffer. The chart
# recolours and relabels itself rather than quietly calling a cost a cushion.
carry_earned = SIGN * carry
PAYS_CARRY   = carry_earned > 0
cushion = abs((FWD - SPOT) / SPOT) * 100.0
c2v = carry / IMPLIED_VOL if IMPLIED_VOL else None
p_sum = sum(s[2] for s in SCENARIOS)
E_R = sum(p * float(total_return(S)) for _, S, p, _ in SCENARIOS)

if abs(p_sum - 1.0) > 1e-9:
    print(f"  ! scenario probabilities sum to {p_sum:.2f}, not 1.00 — "
          f"expected return of {E_R:+.2f}% is not a true expectation")

# --- axis ranges, derived if not pinned -------------------------------------
sc_x = [s[1] for s in SCENARIOS]
x_min = X_MIN if X_MIN is not None else min(sc_x + [SPOT, FWD]) * 0.95
x_max = X_MAX if X_MAX is not None else max(sc_x + [SPOT, FWD]) * 1.05
_yv = total_return(np.linspace(x_min, x_max, 200))
y_min = Y_MIN if Y_MIN is not None else float(_yv.min()) - 8
y_max = Y_MAX if Y_MAX is not None else float(_yv.max()) + 8
x_step = X_STEP if X_STEP else (x_max - x_min) / 6
y_step = Y_STEP if Y_STEP else 10.0

# --- figure ------------------------------------------------------------------
# matplotlib walks FONTS and takes the first installed face. On a machine
# without Frutiger it falls through to Arial silently, which is what you want —
# but it prints a warning per glyph run, so we mute those and report once.
import logging
import warnings
logging.getLogger("matplotlib.font_manager").setLevel(logging.ERROR)
warnings.filterwarnings("ignore", message="findfont")

from matplotlib import font_manager
_installed = {f.name for f in font_manager.fontManager.ttflist}
_used = next((f for f in FONTS if f in _installed), None)
print(f"  font: {_used or 'matplotlib default (none of FONTS installed)'}")

plt.rcParams.update({"font.family": FONTS, "axes.linewidth": 0.8})
fig, ax = plt.subplots(figsize=(FIG_W_IN, FIG_H_IN), dpi=DPI_PNG)

x = np.linspace(x_min, x_max, 1200)     # 1200 points keeps the curve smooth
y = total_return(x)

# 1. Fills carry the meaning before the reader parses a single number.
ax.fill_between(x, 0, y, where=(y > 0), color=C["green"], alpha=FILL_ALPHA, lw=0)
ax.fill_between(x, 0, y, where=(y < 0), color=C["red"],   alpha=FILL_ALPHA, lw=0)
BAND_C = C["green"] if PAYS_CARRY else C["red"]
ax.axvspan(min(SPOT, FWD), max(SPOT, FWD),
           color=BAND_C, alpha=BAND_ALPHA, lw=0)

# 2. Optional decomposition legs, dashed and thin so the total still dominates.
if SHOW_LEGS:
    ax.plot(x, spot_return(x),  color=C["mut"],   lw=1.4, ls="--", zorder=4)
    ax.plot(x, carry_return(x), color=C["amber"], lw=1.4, ls="--", zorder=4)
    ax.text(x_max, float(carry_return(x_max)), "  carry leg",
            fontsize=8, color=C["amber"], va="center")
    ax.text(x_max, float(spot_return(x_max)), "  spot leg",
            fontsize=8, color=C["mut"], va="center")

# 3. One heavy line for the answer.
ax.plot(x, y, color=C[CURVE_COLOUR], lw=CURVE_LW, zorder=5)
ax.axhline(0, color=C["ink"], lw=1.0, zorder=4)

# 4. Reference verticals labelled IN PLACE. A legend would force the reader to
#    look away from the data and back again.
refs = list(EXTRA_REFS)
if AUTO_REFS:
    fmt_x = lambda v: f"{v:,.{X_DECIMALS}f}"
    refs = [(SPOT, f"Spot\n{fmt_x(SPOT)}", "ink", "-"),
            (FWD,  f"Breakeven\n{fmt_x(FWD)}", "red", "--")] + refs
for lvl, lab, ck, ls in refs:
    if not (x_min <= lvl <= x_max):
        continue
    ax.axvline(lvl, color=C[ck], lw=1.2, ls=ls, zorder=3)
    # Labels sit INSIDE the plot at the top. Placing them above the axes
    # collides with the subtitle, and moving the subtitle up to make room
    # just wastes vertical space in the Word column.
    ax.text(lvl, y_max - (y_max - y_min) * 0.02, lab, ha="center", va="top",
            fontsize=8.5, color=C[ck], linespacing=1.4, zorder=7)

# 5. Magnitude stated on an arrow so nobody has to measure it off the axis.
span = y_max - y_min
lab_c = CUSHION_LABEL or (f"{cushion:.1f}% cushion" if PAYS_CARRY
                          else f"{cushion:.1f}% forward premium (a cost)")
_ay = y_min + span * 0.07
ax.annotate("", xy=(FWD, _ay), xytext=(SPOT, _ay),
            arrowprops=dict(arrowstyle="<->", color=BAND_C, lw=1.4))
ax.text((SPOT + FWD) / 2, _ay + span * 0.022, lab_c, ha="center",
        fontsize=9.5, color=BAND_C, fontweight="bold")

# 6. Hollow markers so the curve reads through them; value printed on the label.
for nm, S, p, up in SCENARIOS:
    v = float(total_return(S))
    col = C["green"] if v > 0 else C["red"]
    ax.plot([S], [v], "o", ms=8.5, mfc="white", mec=col, mew=2.4, zorder=6)
    # Offsets scale with the y-range so they behave on a TRY chart (+-45%) as
    # well as a EUR one (+-15%). First value positions the name, second the
    # smaller "level - probability" line, which always sits nearer the marker.
    ABOVE, BELOW = (0.088, 0.050), (-0.112, -0.152)
    d1, d2 = ABOVE if up else BELOW
    # Auto-flip: a label landing on the zero line is unreadable. If either line
    # of the chosen side collides, take the whole other side rather than just
    # negating, which would put the sub-label above the name.
    if min(abs(v + d1 * span), abs(v + d2 * span)) < span * 0.035:
        d1, d2 = BELOW if up else ABOVE
    dy1, dy2 = d1 * span, d2 * span
    ax.text(S, v + dy1, f"{nm}  {v:+.1f}%", ha="center",
            fontsize=9.5, color=col, fontweight="bold")
    ax.text(S, v + dy2, f"{S:,.{X_DECIMALS}f} \u00b7 {p:.0%}",
            ha="center", fontsize=8, color=C["mut"])

# 7. Axes: horizontal gridlines only, three spines removed, no tick marks,
#    signed percentages. This is most of what separates it from a default chart.
ax.set_xlim(x_min, x_max)
ax.set_ylim(y_min, y_max)
ax.set_yticks(np.arange(np.ceil(y_min / y_step) * y_step, y_max + 1e-9, y_step))
ax.set_xticks(np.arange(x_min, x_max + 1e-9, x_step))
ax.yaxis.set_major_formatter(FuncFormatter(lambda v, _: f"{v:+.0f}%"))
ax.xaxis.set_major_formatter(
    FuncFormatter(lambda v, _: f"{v:,.{X_DECIMALS}f}"))
ax.grid(axis="y", color=C["grid"], lw=0.8)
ax.set_axisbelow(True)                       # grid behind the data, always
for s in ("top", "right", "left"):
    ax.spines[s].set_visible(False)
ax.spines["bottom"].set_color(C["grid"])
ax.tick_params(labelsize=8.5, colors=C["mut"], length=0)
ax.set_xlabel(XLABEL, fontsize=9, color=C["mut"], labelpad=6)

# 8. Title left-aligned and stating the finding, not the variables.
ax.set_title(TITLE, fontsize=12, fontweight="bold", color=C["ink"],
             loc="left", pad=26)

sub = SUBTITLE
if sub is None:
    sub = (f"Carry {carry_earned:+.2f}% locked at inception \u00b7 breakeven "
           f"{FWD:,.{X_DECIMALS}f} \u00b7 expected return {E_R:+.1f}%")
    if c2v:
        sub += f" \u00b7 carry/vol {carry_earned / IMPLIED_VOL:.2f}"
ax.text(0, 1.028, sub, transform=ax.transAxes, fontsize=8.5, color=C["mut"])

# --- save --------------------------------------------------------------------
fig.tight_layout()
os.makedirs(OUTDIR, exist_ok=True)
for ext in SAVE:
    path = os.path.join(OUTDIR, f"{OUTNAME}.{ext}")
    fig.savefig(path,
                dpi=DPI_PNG if ext == "png" else None,
                bbox_inches="tight",          # crops dead whitespace so the
                                              # image sits tight in Word
                facecolor="white",            # not transparent — a transparent
                                              # PNG goes grey on some themes
                pad_inches=0.12,
                metadata=None if ext == "png" else {"Creator": "matplotlib"})
    print(f"  wrote {path}")

print(f"\n  carry {carry:+.4f}%   forward {FWD:,.4f}   cushion {cushion:.2f}%"
      f"   E[R] {E_R:+.2f}%")
for nm, S, p, _ in SCENARIOS:
    print(f"    {nm:5s} {S:>10,.{X_DECIMALS}f}  spot {float(spot_return(S)):+7.2f}%"
          f"  carry {float(carry_return(S)):+6.2f}%"
          f"  total {float(total_return(S)):+7.2f}%")


# =============================================================================
#  WORKED CONFIGS FOR OTHER PAIRS
#  Paste over the CONFIG block. Rates are placeholders — pull live from BBG.
# =============================================================================
#
#  MEXICAN PESO — USDMXN, quoted pesos per dollar, same shape as COP
#     PAIR="USDMXN"; BASE_CCY="MXN"; SPOT=17.20
#     RATE_FCY=6.50   # Banxico policy / MXIBTIIE or the 1y TIIE swap
#     RATE_USD=4.02
#     QUOTE="FCY_PER_USD"; DIRECTION="LONG_FCY"
#     X_MIN,X_MAX=15.0,22.0; X_STEP=1.0; X_DECIMALS=2
#     Y_MIN,Y_MAX=-20,20; XLABEL="Terminal USDMXN"
#     SCENARIOS=[("Bull",16.20,0.25,True),("Base",17.80,0.50,False),
#                ("Bear",20.50,0.25,True)]
#
#  BRAZILIAN REAL — USDBRL. Deliverable, so LINEAR against the 1y DI ties well.
#     PAIR="USDBRL"; BASE_CCY="BRL"; SPOT=5.40
#     RATE_FCY=14.00  # Selic, or the 1y DI (ODF pre) for a tradeable number
#     RATE_USD=4.02
#     QUOTE="FCY_PER_USD"; DIRECTION="LONG_FCY"
#     X_MIN,X_MAX=4.50,7.50; X_STEP=0.50; X_DECIMALS=2; X_DECIMALS applies to
#     the scenario labels too. IMPLIED_VOL=15.0
#     Note: Copom is cutting, so the cushion decays over the horizon. Consider
#     running a second chart at the forward-implied end-year Selic to show it.
#
#  TURKISH LIRA — USDTRY. Carry is enormous and so is the tail; pin Y wide.
#     PAIR="USDTRY"; BASE_CCY="TRY"; SPOT=34.0; RATE_FCY=45.0; RATE_USD=4.02
#     QUOTE="FCY_PER_USD"; X_MIN,X_MAX=34.0,70.0; X_STEP=5.0; X_DECIMALS=1
#     Y_MIN,Y_MAX=-45,45; Y_STEP=15; IMPLIED_VOL=25.0
#     The curve is visibly convex over this range — that convexity IS the
#     story, so do not narrow the x-range to flatten it.
#
#  EURO — EURUSD. Quoted the other way round, and EUR is the LOW yielder, so
#  this is a negative-carry long. Set DIRECTION="SHORT_FCY" to chart the
#  funded-in-EUR trade instead.
#     PAIR="EURUSD"; BASE_CCY="EUR"; SPOT=1.0800
#     RATE_FCY=2.00; RATE_USD=4.02
#     QUOTE="USD_PER_FCY"; DIRECTION="LONG_FCY"
#     X_MIN,X_MAX=0.95,1.25; X_STEP=0.05; X_DECIMALS=4
#     Y_MIN,Y_MAX=-15,15; Y_STEP=5
#     CUSHION_LABEL="1.9% forward premium (a cost, not a cushion)"
#     The shaded band now sits to the LEFT of spot and represents ground you
#     must make up, not a buffer. Relabel it or the chart lies.
#
#  SHORTER HORIZON — any pair, 3-month
#     HORIZON_Y=0.25; RATE_FCY = the 3m point, RATE_USD = TSFR3M
#     Cushion scales roughly with h, so it shrinks to about a quarter. Retitle:
#     TITLE="3-month total return ..." — the script does not do it for you.
# =============================================================================
