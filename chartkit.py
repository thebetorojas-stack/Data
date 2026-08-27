# -*- coding: utf-8 -*-
"""
=============================================================================
 chartkit.py  —  house chart module
=============================================================================

 Put this file next to your chart scripts (or anywhere on PYTHONPATH) and
 import it from every chart you build. It holds three things:

   1. The house style  — palette, fonts, axis treatment. Change it here and
                         every chart you have ever written changes with it.
   2. Automatic labels — labels that place themselves and never overlap.
                         This is the part that stops you hand-nudging
                         coordinates on every new chart.
   3. Saving           — SVG + PNG + PDF in one call, sized for Word.

 Requires only matplotlib.

 -------------------------------------------------------------------------
 WHY AUTOMATIC PLACEMENT RATHER THAN MANUAL OFFSETS
 -------------------------------------------------------------------------
 A hand-typed offset is correct for exactly one set of numbers. Move spot,
 add a scenario, switch to a currency with a different axis range, and every
 offset is wrong again. The Labeller measures where the text ACTUALLY lands
 once matplotlib has rendered it, then moves anything that collides. It is
 the same job a person does by eye, done by the machine, every run.

 -------------------------------------------------------------------------
 MINIMAL USE
 -------------------------------------------------------------------------
     import chartkit as ck

     ck.house_style()
     fig, ax = plt.subplots(figsize=(9.75, 5.0))
     ax.plot(x, y, color=ck.C["navy"], lw=2.6)

     lab = ck.Labeller(ax)
     lab.avoid_curve(x, y)                    # keep text off the line
     lab.vline(3152, "Spot\\n3,152", "ink")
     lab.vline(3422, "Breakeven\\n3,422", "red", ls="--")
     lab.point(2950, 16.0, "Bull  +16.0%", "2,950 - 20%", "green")
     lab.resolve()                            # <- call once, after everything

     ck.clean_axes(ax, ypct=True)
     ck.save_all(fig, "my_chart")
=============================================================================
"""

import os
import matplotlib
import matplotlib.pyplot as plt
from matplotlib.ticker import FuncFormatter
from matplotlib.transforms import Bbox

# ============================================================== PALETTE =====
# One accent, one positive, one negative, one grey. Resist adding more.
C = {
    "red":   "#E60000",     # UBS red — the accent
    "green": "#2E7D5B",
    "navy":  "#1F3B63",
    "ink":   "#16202E",
    "mut":   "#6B7A8F",
    "amber": "#C08A2E",
    "grid":  "#E3E8EF",
}

FONTS = ["Frutiger 45 Light", "Frutiger", "Arial", "Helvetica", "DejaVu Sans"]


def house_style(fonts=None, quiet=True):
    """Set the house rcParams. Call once, before creating the figure."""
    if quiet:
        import logging, warnings
        logging.getLogger("matplotlib.font_manager").setLevel(logging.ERROR)
        warnings.filterwarnings("ignore", message="findfont")
    fonts = fonts or FONTS
    from matplotlib import font_manager
    have = {f.name for f in font_manager.fontManager.ttflist}
    used = next((f for f in fonts if f in have), None)
    plt.rcParams.update({
        "font.family": fonts,
        "axes.linewidth": 0.8,
        "svg.fonttype": "none",   # keep text as text in the SVG, not outlines,
                                  # so Word can re-render it crisply
    })
    return used


# ============================================================ AXIS STYLE ====
def clean_axes(ax, xfmt=None, yfmt=None, ypct=False, xcomma=True,
               xdecimals=0, xlabel=None, ylabel=None, grid="y"):
    """Strip a matplotlib default down to the house treatment.

    Horizontal gridlines only, three of four spines removed, no tick marks,
    grid behind the data. This one function is most of what separates a house
    chart from a stock one.
    """
    if grid:
        ax.grid(axis=grid, color=C["grid"], lw=0.8)
    ax.set_axisbelow(True)
    for s in ("top", "right", "left"):
        ax.spines[s].set_visible(False)
    ax.spines["bottom"].set_color(C["grid"])
    ax.tick_params(labelsize=8.5, colors=C["mut"], length=0)
    if ypct:
        ax.yaxis.set_major_formatter(FuncFormatter(lambda v, _: f"{v:+.0f}%"))
    if yfmt:
        ax.yaxis.set_major_formatter(FuncFormatter(yfmt))
    if xcomma and not xfmt:
        ax.xaxis.set_major_formatter(
            FuncFormatter(lambda v, _: f"{v:,.{xdecimals}f}"))
    if xfmt:
        ax.xaxis.set_major_formatter(FuncFormatter(xfmt))
    if xlabel:
        ax.set_xlabel(xlabel, fontsize=9, color=C["mut"], labelpad=6)
    if ylabel:
        ax.set_ylabel(ylabel, fontsize=9, color=C["mut"], labelpad=6)


def titles(ax, title, subtitle=None):
    """Left-aligned title stating the finding, muted subtitle beneath."""
    ax.set_title(title, fontsize=12, fontweight="bold", color=C["ink"],
                 loc="left", pad=26)
    if subtitle:
        ax.text(0, 1.028, subtitle, transform=ax.transAxes,
                fontsize=8.5, color=C["mut"])


# ======================================================= LABEL PLACEMENT ====
class Labeller:
    """Collision-free text placement.

    You register labels, then call resolve() once. resolve() renders the
    figure, measures every label's true pixel footprint, and moves any that
    overlap each other, the plotted line, or the edge of the axes.

    Two kinds of label:
      vline(x, text, ...)  — a vertical reference line with text near the top
      point(x, y, ...)     — text attached to a marker

    Candidate positions are tried in order. For a vline label: centred on the
    line, then right of it, then left of it; then the whole sequence again one
    text-height lower, and so on. For a point label: above the marker, then
    below, then progressively further out. The first candidate that fits wins,
    so the common case still looks hand-placed.
    """

    def __init__(self, ax, pad_px=3.0, max_rows=16):
        self.ax = ax
        self.pad = pad_px          # clear space demanded around each label
        self.max_rows = max_rows   # how many vertical rows to try before
                                   # giving up and accepting the last attempt.
                                   # Raise it if you have many crowded labels;
                                   # it only costs a few milliseconds.
        self._v = []
        self._p = []
        self._obstacles = []       # extra boxes text must avoid

    # ---------------------------------------------------------- register --
    def vline(self, x, text, colour="mut", ls="-", lw=1.2, draw_line=True,
              fontsize=8.5, anchor=0.98):
        """A vertical reference line plus its label.

        anchor: where the label starts vertically, 1.0 = top of plot.
        """
        self._v.append(dict(x=x, text=text, colour=colour, ls=ls, lw=lw,
                            draw_line=draw_line, fs=fontsize, anchor=anchor))

    def point(self, x, y, text, sub=None, colour="ink", fontsize=9.5,
              subsize=8, marker=True, ms=8.5):
        """Text attached to a data point, optionally with a marker drawn."""
        self._p.append(dict(x=x, y=y, text=text, sub=sub, colour=colour,
                            fs=fontsize, ss=subsize, marker=marker, ms=ms))

    def avoid_curve(self, xs, ys, every=6):
        """Treat a plotted line as an obstacle so text does not sit on it.

        `every` thins the sample — every 6th point is plenty and keeps
        resolve() fast even on a 1200-point curve.
        """
        self._obstacles.append(("curve", list(xs)[::every], list(ys)[::every]))

    def avoid_box(self, x0, y0, x1, y1):
        """Reserve a rectangle in DATA coordinates (e.g. where a legend sits)."""
        self._obstacles.append(("box", (x0, y0), (x1, y1)))

    def avoid_hline(self, y, thickness=0.004):
        """Keep text off a horizontal reference line such as zero.

        thickness is a fraction of the y-range. Convenience wrapper around
        avoid_box across the full width of the axes.
        """
        x0, x1 = self.ax.get_xlim()
        y0, y1 = self.ax.get_ylim()
        h = (y1 - y0) * thickness
        self._obstacles.append(("box", (x0, y - h), (x1, y + h)))

    # ------------------------------------------------------------ resolve --
    def resolve(self):
        ax, fig = self.ax, self.ax.figure
        fig.canvas.draw()                      # must render before measuring
        rend = fig.canvas.get_renderer()
        t = ax.transData
        axbox = ax.get_window_extent(rend)
        x0, x1 = ax.get_xlim()
        y0, y1 = ax.get_ylim()
        xspan, yspan = x1 - x0, y1 - y0

        taken = self._obstacle_boxes(rend)

        # --- draw the lines and markers first; they are not moveable -------
        # They also become obstacles, which is what pushes a label off its own
        # vertical line and out from under its own marker.
        for v in self._v:
            if v["draw_line"] and x0 <= v["x"] <= x1:
                ax.axvline(v["x"], color=C[v["colour"]], lw=v["lw"],
                           ls=v["ls"], zorder=3)
            px = t.transform((v["x"], y0))[0]
            taken.append(Bbox.from_extents(px - 1.5, axbox.y0,
                                           px + 1.5, axbox.y1))
        for p in self._p:
            if p["marker"]:
                ax.plot([p["x"]], [p["y"]], "o", ms=p["ms"], mfc="white",
                        mec=C[p["colour"]], mew=2.4, zorder=6)
            px, py = t.transform((p["x"], p["y"]))
            r = p["ms"] * fig.dpi / 72.0 / 2 + 1
            taken.append(Bbox.from_extents(px - r, py - r, px + r, py + r))

        # --- vertical-line labels, left to right ---------------------------
        for v in sorted(self._v, key=lambda d: d["x"]):
            if not (x0 <= v["x"] <= x1):
                continue
            txt = ax.text(v["x"], y0 + yspan * v["anchor"], v["text"],
                          ha="center", va="top", fontsize=v["fs"],
                          color=C[v["colour"]], linespacing=1.4, zorder=7)
            h = txt.get_window_extent(rend).height
            dx = xspan * 0.012                 # one character, roughly
            dy_row = (h / axbox.height) * yspan * 1.10
            cands = []
            for row in range(self.max_rows):
                yy = y0 + yspan * v["anchor"] - row * dy_row
                for xx, ha in ((v["x"], "center"),
                               (v["x"] + dx, "left"),
                               (v["x"] - dx, "right"),
                               (v["x"] + dx * 3, "left"),
                               (v["x"] - dx * 3, "right")):
                    cands.append(((xx, yy), ha))
            taken.append(self._best(txt, cands, taken, axbox, rend))

        # --- point labels ---------------------------------------------------
        for p in self._p:
            main = ax.text(p["x"], p["y"], p["text"], ha="center",
                           fontsize=p["fs"], color=C[p["colour"]],
                           fontweight="bold", zorder=7)
            sub = (ax.text(p["x"], p["y"], p["sub"], ha="center",
                           fontsize=p["ss"], color=C["mut"], zorder=7)
                   if p["sub"] else None)
            hm = main.get_window_extent(rend).height / axbox.height * yspan
            hs = (sub.get_window_extent(rend).height / axbox.height * yspan
                  if sub else 0.0)
            gap = yspan * 0.030                # clearance from the marker
            step = yspan * 0.035               # how far each retry moves out
            # Label width in data units, so sideways candidates scale with the
            # text rather than with the axis. On a steep curve, stepping aside
            # clears the line far sooner than stepping further up or down —
            # without this the label ends up marooned from its own marker.
            wm = main.get_window_extent(rend).width / axbox.width * xspan
            cands = []
            for k in range(self.max_rows):
                out = gap + k * step
                for dx in (0.0, wm * 0.62, -wm * 0.62, wm * 1.15, -wm * 1.15):
                    # above: sub nearer the marker, name above it
                    cands.append((dx, p["y"] + out + hs * 1.05, p["y"] + out))
                    # below: name nearer, sub beneath it
                    cands.append((dx, p["y"] - out - hm * 1.05,
                                      p["y"] - out - hm * 1.05 - hs * 1.05))
            best, best_score = None, None
            for dx, ym, ys in cands:
                main.set_position((p["x"] + dx, ym))
                if sub:
                    sub.set_position((p["x"] + dx, ys))
                bm = main.get_window_extent(rend)
                sc = self._score(bm, taken, axbox)
                # Among candidates that all fit, prefer the one nearest its
                # own marker — otherwise a label can end up cleanly placed but
                # marooned halfway across the chart from the point it labels.
                # The weight is small relative to overlap area in square
                # pixels, so proximity only ever breaks ties; it never
                # overrides a real collision.
                sc += (abs(dx) / xspan +
                       abs(ym - p["y"]) / yspan) * 120.0
                if sub:
                    bs = sub.get_window_extent(rend)
                    sc += self._score(bs, taken + [bm], axbox)
                if best_score is None or sc < best_score:
                    best, best_score = (dx, ym, ys), sc
            bdx, bym, bys = best
            main.set_position((p["x"] + bdx, bym))
            taken.append(main.get_window_extent(rend))
            if sub:
                sub.set_position((p["x"] + bdx, bys))
                taken.append(sub.get_window_extent(rend))

    # ------------------------------------------------------------ helpers --
    def _score(self, bb, taken, axbox):
        """Badness of a candidate position, in square pixels. 0 = perfect.

        Scoring rather than a pass/fail test matters: when a chart is genuinely
        too crowded for any position to be clean, we still want the LEAST bad
        one, not whichever attempt happened to be tried last. Leaving the axes
        is weighted heavily, so a label will always stay inside the plot even
        if that means clipping a gridline.
        """
        pad = self.pad
        g = Bbox.from_extents(bb.x0 - pad, bb.y0 - pad,
                              bb.x1 + pad, bb.y1 + pad)
        outside = (max(0, axbox.x0 - g.x0) + max(0, g.x1 - axbox.x1) +
                   max(0, axbox.y0 - g.y0) + max(0, g.y1 - axbox.y1))
        overlap = 0.0
        for b in taken:
            ix = min(g.x1, b.x1) - max(g.x0, b.x0)
            iy = min(g.y1, b.y1) - max(g.y0, b.y0)
            if ix > 0 and iy > 0:
                overlap += ix * iy
        return outside * 1e4 + overlap

    def _best(self, txt, cands, taken, axbox, rend):
        """Try every candidate, keep the lowest-scoring one, return its bbox."""
        best, best_score = None, None
        for pos, ha in cands:
            txt.set_position(pos)
            txt.set_ha(ha)
            sc = self._score(txt.get_window_extent(rend), taken, axbox)
            if best_score is None or sc < best_score:
                best, best_score = (pos, ha), sc
            if sc == 0:
                break
        txt.set_position(best[0])
        txt.set_ha(best[1])
        return txt.get_window_extent(rend)

    def _obstacle_boxes(self, rend):
        """Turn registered curves and boxes into display-coordinate rectangles."""
        t = self.ax.transData
        out = []
        for kind, *rest in self._obstacles:
            if kind == "curve":
                xs, ys = rest
                pts = t.transform(list(zip(xs, ys)))
                for (px, py) in pts:
                    out.append(Bbox.from_extents(px - 2, py - 2, px + 2, py + 2))
            else:
                (ax0, ay0), (ax1, ay1) = rest
                (px0, py0), (px1, py1) = t.transform([(ax0, ay0), (ax1, ay1)])
                out.append(Bbox.from_extents(min(px0, px1), min(py0, py1),
                                             max(px0, px1), max(py0, py1)))
        return out


# ================================================================= SAVE =====
def save_all(fig, stem, outdir=".", dpi=600, formats=("svg", "png", "pdf")):
    """Write the chart in every format, sized and cropped for Word.

    Insert the SVG into Word (Insert > Pictures > This Device): it stays true
    vector and never pixelates. The PNG at 600 dpi is the fallback for older
    builds. Never enlarge the PNG inside Word — change figsize and re-run.
    """
    fig.tight_layout()
    os.makedirs(outdir, exist_ok=True)
    paths = []
    for ext in formats:
        path = os.path.join(outdir, f"{stem}.{ext}")
        fig.savefig(path, dpi=dpi if ext == "png" else None,
                    bbox_inches="tight", facecolor="white", pad_inches=0.12)
        paths.append(path)
        print(f"  wrote {path}")
    return paths
