# -*- coding: utf-8 -*-
"""
=============================================================================
 FX CARRY PAYOFF CHART
=============================================================================

 Draws the 12-month total return on a forward / NDF position against terminal
 spot, with the breakeven cushion shaded and scenario points marked. The whole
 point of the chart is that carry is fixed the day you put the trade on, so
 terminal spot is the only variable and the entire payoff collapses to one
 curve.

 FILES
 -----
   chartkit.py          must sit in the same folder. Holds the house style and
                        the automatic label placement.
   fx_carry_chart.py    this file. Holds the trade and the scenarios.

 RUN
 ---
   pip install matplotlib          the only dependency
   python fx_carry_chart.py

 OUTPUT
 ------
   <OUTNAME>.svg   INSERT THIS INTO WORD (Insert > Pictures > This Device).
                   True vector: sharp at any zoom, at any print resolution,
                   and through a PDF distill. There is no resolution to run
                   out of.
   <OUTNAME>.png   600 dpi fallback for older Word or locked-down PowerPoint.
                   Never enlarge it inside Word — change FIG_W_IN and re-run,
                   so the pixels are generated rather than interpolated.
   <OUTNAME>.pdf   vector, for LaTeX or for handing to design.

 TO MOVE TO ANOTHER CURRENCY
 ---------------------------
 Edit the CONFIG block below. Nothing beneath the ENGINE line needs touching.
 The two settings that actually matter are QUOTE and DIRECTION — get those
 right and every sign in the chart follows automatically. Worked configs for
 MXN, BRL, TRY and EUR are at the bottom of this file.

 There are no label coordinates anywhere in here. chartkit measures where each
 label really lands and moves anything that collides, so changing spot, adding
 a scenario, or switching to a currency that trades at 17 instead of 3,152
 needs no repositioning by hand.
=============================================================================
"""

import matplotlib
matplotlib.use("Agg")                    # write files, never open a window
import matplotlib.pyplot as plt
import numpy as np
import chartkit as ck


# =============================================================== CONFIG =====

# --- the trade ---------------------------------------------------------------
PAIR      = "USDCOP"       # cosmetic, used in the auto x-axis label
LOCAL_CCY = "COP"          # cosmetic, used in the auto title
SPOT      = 3152.0         # current spot, in the units named by QUOTE
RATE_LOC  = 12.5911        # % p.a. local rate.    BBG: CLNI12M BGN Curncy
RATE_USD  = 4.02           # % p.a. funding leg.   BBG: USOSFR1 BGN Curncy
                           #   Never LIBOR: ICE ceased synthetic USD
                           #   publication on 30 Sep 2024. Use SOFR OIS
                           #   (USOSFR1) or Term SOFR (TSFR3M / TSFR1M).
HORIZON_Y = 1.0            # years to the forward date. 1.0 = 12m, 0.25 = 3m.

QUOTE = "LOCAL_PER_USD"
#   "LOCAL_PER_USD"  units of local currency per 1 USD. A HIGHER number means
#                    a WEAKER local currency. COP MXN BRL CLP PEN TRY ZAR JPY
#                    INR IDR PHP.
#   "USD_PER_LOCAL"  USD per 1 unit of local. A HIGHER number means a STRONGER
#                    local currency. EUR GBP AUD NZD.

DIRECTION = "LONG_LOCAL"
#   "LONG_LOCAL"   long the local currency, funded in USD. The normal carry
#                  trade when RATE_LOC is the higher rate.
#   "SHORT_LOCAL"  the reverse. Flips every sign in the chart.

FWD_METHOD = "LINEAR"
#   "LINEAR"  F = S x (1 + carry x h). The desk convention of simply adding the
#             quoted differential. Use this to tie out to a sheet built that way.
#   "EXACT"   F = S x (1 + r_loc x h) / (1 + r_usd x h). Covered interest parity.
#             The two diverge as rates rise — at 12.6% against 4.0% over 12m
#             they differ by about 10 big figures on USDCOP. Pick one
#             deliberately and say which one in the footnote.
#   Neither isolates the cross-currency basis. For a deliverable NDF P&L, take
#   the outright straight off the quoted NDF points via FWD_OVERRIDE.
FWD_OVERRIDE = None        # e.g. 3422.0 to hard-set the outright. None = compute.

IMPLIED_VOL = 13.0         # % p.a. Set None to drop the carry/vol figure from
                           # the subtitle. Below ~0.5 the trade is thin; above
                           # ~1.0 it is genuinely attractive.

# --- scenarios ---------------------------------------------------------------
# (name, terminal spot, probability)
# Probabilities feed the expected return in the subtitle. They need not sum to
# 1 — if they don't, the script says so rather than silently normalising.
# There is no "label above or below" setting: chartkit works that out.
SCENARIOS = [
    ("Bull", 2950.0, 0.20),
    ("Base", 3250.0, 0.55),
    ("Bear", 3950.0, 0.25),
]

# --- extra reference lines ---------------------------------------------------
# (level, label text, colour key, linestyle). Use \n for a second line.
# Spot and breakeven are added automatically. Colour keys: ink red green navy
# mut amber. Set to [] for none.
EXTRA_REFS = [
    (3700.0, "Start-2026\n~3,700", "mut", ":"),
]

# --- axes --------------------------------------------------------------------
# Any of these may be None to derive it from the data. Pinning them is better
# for a series of charts you want directly comparable.
X_MIN, X_MAX = 2800.0, 4300.0
X_STEP       = 250.0       # gap between x tick labels
Y_MIN, Y_MAX = -24.0, 26.0 # percent
Y_STEP       = 10.0
X_DECIMALS   = 0           # 0 for COP CLP JPY · 2 for MXN BRL · 4 for EURUSD

# --- text --------------------------------------------------------------------
# All None = built automatically from the config above, which is usually right.
TITLE         = None
SUBTITLE      = None
XLABEL        = None
CUSHION_LABEL = None

# --- output ------------------------------------------------------------------
OUTDIR   = "."
OUTNAME  = "fx_carry_chart"
FIG_W_IN = 9.75            # inches. Match your Word column: US Letter with 1"
                           # margins is 6.5"; A4 with 2.5cm is 6.3". Drawing
                           # WIDER than the column and letting Word scale it
                           # down is deliberate — it makes every line and glyph
                           # finer. 9.75" into 6.5" is a 0.67x reduction.
FIG_H_IN = 5.00
DPI_PNG  = 600             # 600 is print grade. 300 is the floor. Beyond 600
                           # the file bloats with no visible gain.
FORMATS  = ("svg", "png", "pdf")

# --- style -------------------------------------------------------------------
CURVE_COLOUR = "navy"
CURVE_LW     = 2.6         # one heavy line for the answer, everything thinner
FILL_ALPHA   = 0.09        # tint under the curve
BAND_ALPHA   = 0.12        # the cushion band
SHOW_LEGS    = False       # True adds dashed spot-return and carry-return
                           # lines. Useful for a desk audience, noise for a
                           # client one.


# =============================================================== ENGINE =====
# Nothing below needs editing to swap currency. It is commented anyway.

C = ck.C
h = HORIZON_Y
carry = RATE_LOC - RATE_USD               # % p.a., local minus funding

# --- forward outright --------------------------------------------------------
# Quoted as local-per-USD, a higher-yielding local currency trades at a HIGHER
# forward — it depreciates in the forward, and that is exactly what hands you
# the cushion. Quoted the other way round the arithmetic inverts.
if FWD_OVERRIDE is not None:
    FWD = float(FWD_OVERRIDE)
elif QUOTE == "LOCAL_PER_USD":
    FWD = (SPOT * (1 + carry / 100 * h) if FWD_METHOD == "LINEAR"
           else SPOT * (1 + RATE_LOC / 100 * h) / (1 + RATE_USD / 100 * h))
elif QUOTE == "USD_PER_LOCAL":
    FWD = (SPOT * (1 - carry / 100 * h) if FWD_METHOD == "LINEAR"
           else SPOT * (1 + RATE_USD / 100 * h) / (1 + RATE_LOC / 100 * h))
else:
    raise ValueError("QUOTE must be 'LOCAL_PER_USD' or 'USD_PER_LOCAL'")

SIGN = 1.0 if DIRECTION == "LONG_LOCAL" else -1.0


def total_return(S):
    """Total return in % on the funding currency's notional.

    LOCAL_PER_USD: you sold USD forward at FWD and buy it back at S, so the
      P&L per unit of USD notional is (FWD - S) / S. Dividing by S and not by
      FWD is what makes the loss convex, and it is precisely the term the naive
      additive decomposition carry + (S0/S - 1) gets wrong.
    USD_PER_LOCAL: you bought the local currency forward at FWD and sell it at
      S, giving (S - FWD) / FWD.
    """
    S = np.asarray(S, dtype=float)
    if QUOTE == "LOCAL_PER_USD":
        return SIGN * (FWD - S) / S * 100.0
    return SIGN * (S - FWD) / FWD * 100.0


def spot_return(S):
    """The pure currency move, carry held aside."""
    S = np.asarray(S, dtype=float)
    if QUOTE == "LOCAL_PER_USD":
        return SIGN * (SPOT / S - 1) * 100.0
    return SIGN * (S / SPOT - 1) * 100.0


def carry_return(S):
    """Carry actually earned. NOT a constant.

    The coupon is fixed in local currency but you realise it at the TERMINAL
    exchange rate, so its USD value scales by SPOT/S. If the currency halves,
    so does the carry. Drawing carry as a flat line is the commonest error in
    these decompositions and it always flatters the bear.
    """
    return total_return(S) - spot_return(S)


# Carry you actually EARN, after DIRECTION. Negative means the forward works
# against you — long EUR against USD at current rates pays a premium rather
# than collecting one. The chart recolours and relabels itself in that case
# rather than quietly calling a cost a cushion.
carry_earned = SIGN * carry
PAYS_CARRY   = carry_earned > 0
cushion      = abs((FWD - SPOT) / SPOT) * 100.0

p_sum = sum(p for _, _, p in SCENARIOS)
E_R = sum(p * float(total_return(S)) for _, S, p in SCENARIOS)
if abs(p_sum - 1.0) > 1e-9:
    print(f"  ! probabilities sum to {p_sum:.2f}, not 1.00 — the "
          f"{E_R:+.2f}% figure is not a true expectation")

# --- axis ranges, derived where not pinned ----------------------------------
sc_x = [S for _, S, _ in SCENARIOS]
x_min = X_MIN if X_MIN is not None else min(sc_x + [SPOT, FWD]) * 0.95
x_max = X_MAX if X_MAX is not None else max(sc_x + [SPOT, FWD]) * 1.05
_probe = total_return(np.linspace(x_min, x_max, 200))
y_min = Y_MIN if Y_MIN is not None else float(_probe.min()) - 8
y_max = Y_MAX if Y_MAX is not None else float(_probe.max()) + 8
x_step = X_STEP or (x_max - x_min) / 6
y_step = Y_STEP or 10.0
span = y_max - y_min

# --- figure ------------------------------------------------------------------
used_font = ck.house_style()
print(f"  font: {used_font or 'matplotlib default'}")
fig, ax = plt.subplots(figsize=(FIG_W_IN, FIG_H_IN), dpi=DPI_PNG)

x = np.linspace(x_min, x_max, 1200)       # 1200 points keeps the curve smooth
y = total_return(x)

# 1. Tints carry the meaning before the reader parses a single number. Keep
#    them faint — shading should register, not shout.
ax.fill_between(x, 0, y, where=(y > 0), color=C["green"], alpha=FILL_ALPHA, lw=0)
ax.fill_between(x, 0, y, where=(y < 0), color=C["red"],   alpha=FILL_ALPHA, lw=0)
BAND_C = C["green"] if PAYS_CARRY else C["red"]
ax.axvspan(min(SPOT, FWD), max(SPOT, FWD), color=BAND_C, alpha=BAND_ALPHA, lw=0)

# 2. Optional decomposition legs, thin and dashed so the total still dominates.
if SHOW_LEGS:
    ax.plot(x, spot_return(x),  color=C["mut"],   lw=1.4, ls="--", zorder=4)
    ax.plot(x, carry_return(x), color=C["amber"], lw=1.4, ls="--", zorder=4)

# 3. One heavy line for the answer, plus the zero axis.
ax.plot(x, y, color=C[CURVE_COLOUR], lw=CURVE_LW, zorder=5)
ax.axhline(0, color=C["ink"], lw=1.0, zorder=4)

ax.set_xlim(x_min, x_max)
ax.set_ylim(y_min, y_max)
ax.set_yticks(np.arange(np.ceil(y_min / y_step) * y_step, y_max + 1e-9, y_step))
ax.set_xticks(np.arange(x_min, x_max + 1e-9, x_step))

# 4. The cushion arrow. Drawn by hand because it spans two x-values rather than
#    marking one, then reserved as an obstacle so no label lands on top of it.
cush_y = y_min + span * 0.07
ax.annotate("", xy=(FWD, cush_y), xytext=(SPOT, cush_y),
            arrowprops=dict(arrowstyle="<->", color=BAND_C, lw=1.4))
cush_lab = CUSHION_LABEL or (
    f"{cushion:.1f}% cushion" if PAYS_CARRY
    else f"{cushion:.1f}% forward premium (a cost)")
ax.text((SPOT + FWD) / 2, cush_y + span * 0.022, cush_lab, ha="center",
        fontsize=9.5, color=BAND_C, fontweight="bold", zorder=7)

# 5. Labels. Register them and let chartkit place them — no coordinates here.
L = ck.Labeller(ax)
L.avoid_curve(x, y)                        # keep text off the payoff line
L.avoid_hline(0)                           # and off the zero axis
L.avoid_box(min(SPOT, FWD), cush_y - span * 0.03,
            max(SPOT, FWD), cush_y + span * 0.07)      # the cushion arrow

fx = lambda v: f"{v:,.{X_DECIMALS}f}"
L.vline(SPOT, f"Spot\n{fx(SPOT)}", "ink")
L.vline(FWD,  f"Breakeven\n{fx(FWD)}", "red", ls="--")
for lvl, lab, col, ls in EXTRA_REFS:
    L.vline(lvl, lab, col, ls=ls)

for nm, S, p in SCENARIOS:
    v = float(total_return(S))
    L.point(S, v, f"{nm}  {v:+.1f}%", f"{fx(S)} \u00b7 {p:.0%}",
            "green" if v > 0 else "red")

L.resolve()                                # one call, after every registration

# 6. House axis treatment and titles.
ck.clean_axes(ax, ypct=True, xdecimals=X_DECIMALS,
              xlabel=XLABEL or f"Terminal {PAIR}")

months = int(round(h * 12))
title = TITLE or (f"{months}-month total return on a "
                  f"{'long' if SIGN > 0 else 'short'}-{LOCAL_CCY} "
                  f"forward, by terminal {PAIR}")
sub = SUBTITLE
if sub is None:
    sub = (f"Carry {carry_earned:+.2f}% locked at inception \u00b7 breakeven "
           f"{fx(FWD)} \u00b7 expected return {E_R:+.1f}%")
    if IMPLIED_VOL:
        sub += f" \u00b7 carry/vol {carry_earned / IMPLIED_VOL:.2f}"
ck.titles(ax, title, sub)

# 7. Save.
ck.save_all(fig, OUTNAME, outdir=OUTDIR, dpi=DPI_PNG, formats=FORMATS)

print(f"\n  carry {carry_earned:+.4f}%   forward {FWD:,.4f}   "
      f"cushion {cushion:.2f}%   E[R] {E_R:+.2f}%")
for nm, S, p in SCENARIOS:
    print(f"    {nm:5s} {S:>10,.{X_DECIMALS}f}  "
          f"spot {float(spot_return(S)):+7.2f}%  "
          f"carry {float(carry_return(S)):+6.2f}%  "
          f"total {float(total_return(S)):+7.2f}%")


# =============================================================================
#  WORKED CONFIGS FOR OTHER PAIRS
#  Paste over the matching lines in CONFIG. Rates are placeholders — pull live.
# =============================================================================
#
#  MEXICAN PESO — USDMXN. Same shape as COP, just smaller numbers.
#     PAIR="USDMXN"; LOCAL_CCY="MXN"; SPOT=17.20
#     RATE_LOC=6.50            # Banxico policy, or the 1y TIIE swap
#     RATE_USD=4.02
#     QUOTE="LOCAL_PER_USD"; DIRECTION="LONG_LOCAL"
#     X_MIN,X_MAX=15.0,22.0; X_STEP=1.0; X_DECIMALS=2
#     Y_MIN,Y_MAX=-20.0,20.0; Y_STEP=5.0; IMPLIED_VOL=11.0
#     SCENARIOS=[("Bull",16.20,0.25),("Base",17.80,0.50),("Bear",20.50,0.25)]
#     EXTRA_REFS=[]
#
#  BRAZILIAN REAL — USDBRL. Deliverable, so LINEAR against the 1y DI ties well.
#     PAIR="USDBRL"; LOCAL_CCY="BRL"; SPOT=5.40
#     RATE_LOC=14.00           # Selic, or 1y DI (ODF pre) for a tradeable rate
#     RATE_USD=4.02
#     QUOTE="LOCAL_PER_USD"; DIRECTION="LONG_LOCAL"
#     X_MIN,X_MAX=4.50,7.50; X_STEP=0.50; X_DECIMALS=2
#     Y_MIN,Y_MAX=-30.0,30.0; Y_STEP=10.0; IMPLIED_VOL=15.0
#     SCENARIOS=[("Bull",4.90,0.25),("Base",5.70,0.50),("Bear",6.80,0.25)]
#     Copom is cutting, so unlike COP the cushion decays across the horizon.
#     Worth a second chart run at the forward-implied end-year Selic to show it.
#
#  TURKISH LIRA — USDTRY. Huge carry, huge tail. Pin Y wide.
#     PAIR="USDTRY"; LOCAL_CCY="TRY"; SPOT=34.0
#     RATE_LOC=45.00; RATE_USD=4.02
#     QUOTE="LOCAL_PER_USD"; DIRECTION="LONG_LOCAL"
#     X_MIN,X_MAX=34.0,70.0; X_STEP=5.0; X_DECIMALS=1
#     Y_MIN,Y_MAX=-45.0,45.0; Y_STEP=15.0; IMPLIED_VOL=25.0
#     The curve is visibly convex over this range and that convexity IS the
#     story, so resist narrowing the x-range to flatten it.
#
#  EURO — EURUSD. Quoted the other way up, and EUR is the LOW yielder, so this
#  is a negative-carry long. The band turns red and relabels itself as a cost.
#     PAIR="EURUSD"; LOCAL_CCY="EUR"; SPOT=1.0800
#     RATE_LOC=2.00; RATE_USD=4.02
#     QUOTE="USD_PER_LOCAL"; DIRECTION="LONG_LOCAL"
#     X_MIN,X_MAX=0.95,1.25; X_STEP=0.05; X_DECIMALS=4
#     Y_MIN,Y_MAX=-15.0,15.0; Y_STEP=5.0; IMPLIED_VOL=8.0
#     SCENARIOS=[("Bull",1.1800,0.30),("Base",1.1000,0.40),("Bear",1.0200,0.30)]
#     EXTRA_REFS=[]
#
#  SHORTER HORIZON — any pair, 3 months
#     HORIZON_Y=0.25; RATE_LOC = the 3m point; RATE_USD = TSFR3M
#     The cushion scales roughly with h, so it shrinks to about a quarter. The
#     title updates itself to "3-month"; nothing else needs changing.
# =============================================================================
