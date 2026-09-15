#!/usr/bin/env python3
"""
bond_ladder.py — 1y-10y EM bond ladders from the EM Bond List universe
=====================================================================

Builds a maturity ladder (one rung per year, 1y ... 10y by default) out of
the SAME eligible universe the weekly EM Bond List publishes. It imports
GEMData from gem_report_builder_v3.py, so eligibility, issuer names, ratings,
IG/HY placement, house view and Restrictions can never drift from the list.

Quick start (Spyder): edit the CONFIG block below, press F5.
Command line:         python bond_ladder.py                 (uses CONFIG)
                      python bond_ladder.py --grade HY
                      python bond_ladder.py --grade IG --spread 100
                      python bond_ladder.py --grade ANY --pick price --region "Latin America"
                      python bond_ladder.py --universe sovereign --per-rung 2

The four "flavours" you asked for, all driven by GRADE / MIN_SPREAD_BP / PICK:
    pure IG ladder ................ GRADE = 'IG'
    pure HY ladder ................ GRADE = 'HY'
    anything >= UST + 100bp ....... GRADE = 'ANY',  MIN_SPREAD_BP = 100   (see section 4:
                                    no Treasury data needed — the EMBI floor does it)
    max yield per rung ............ PICK = 'yield'   (ties -> lowest price)
    lowest price per rung ......... PICK = 'price'   (ties -> highest yield)
    best of both .................. PICK = 'both'    (rank on yield + rank on price)

Inputs:  data/current/*.txt (the 7 weekly feed files). Prev-week, PRIIPS and
         legal-exclusion files are picked up automatically if present, like
         run_weekly.py does, but are optional here.
Output:  outputs/Bond_Ladder_<GRADE>_<PICK>[_<spread>bp].xlsx  +  .pdf (one page, for FAs)
           • Ladder      — the picked bond per rung + summary
           • Candidates  — every eligible bond per rung, ranked (see who lost)
           • Settings    — the CONFIG + the EM curve used, for the audit trail
"""

import argparse
import glob
import os
import sys
from datetime import date

HERE = os.path.dirname(os.path.abspath(__file__))
os.chdir(HERE)
if HERE not in sys.path:
    sys.path.insert(0, HERE)

from gem_report_builder_v3 import (GEMData, is_subordinated_bond,   # noqa: E402
                                   parse_rating, rating_tier,      # business logic lives there
                                   register_fonts, DEFAULT_LOGO_PATH, INDICATIVE_MARK,
                                   UBS_DARK, UBS_MID, UBS_HEADER_BG, UBS_RULE, UBS_LIGHT)

# ══════════════════════════════════════════════════════════════════════════════
# CONFIG — the only block you normally touch
# ══════════════════════════════════════════════════════════════════════════════

# ---- 1. Flavour of ladder --------------------------------------------------
GRADE          = 'IG'        # 'IG' | 'HY' | 'ANY'   (uses the list's own IG/HY placement)
MIN_SPREAD_BP  = None        # e.g. 100 -> only bonds AT LEAST 100bp over UST (section 4). None = off
PICK           = 'yield'     # 'yield' | 'price' | 'both'   (how the winner per rung is chosen)
MIN_RATING     = None        # optional floor on the WORST of S&P/Moody's, e.g. 'BB-' or 'Baa3'. None = off

# ---- 2. Ladder shape -------------------------------------------------------
LADDER_YEARS   = list(range(1, 11))   # rungs: 1y, 2y, ... 10y
RUNG_WINDOW    = 0.5                  # a 5y rung takes bonds maturing 4.5y-5.5y from today
RUNG_WINDOW_FALLBACK = 1.0            # if a rung is empty, widen once to +/- this. None = don't
BONDS_PER_RUNG = 1                    # >1 gives you alternatives on every rung
ONE_BOND_PER_ISSUER = True            # no issuer repeated across the whole ladder
MAX_PER_COUNTRY = None                # e.g. 3 -> at most 3 rungs from the same country. None = off
ALLOW_REPEAT_IF_NEEDED = True         # if a rung would be EMPTY because of the two limits above,
                                      # fill it with a repeat anyway (flagged in the Note column)

# ---- 3. Universe -----------------------------------------------------------
CURRENCY       = 'USD'                # 'USD' | 'EUR' | ... | 'ANY'
UNIVERSE       = 'all'                # 'sovereign' (SOV+SUPRA) | 'corporate' (FIN+CORP) | 'all'
REGION         = None                 # None | 'Latin America' | 'Asia' | 'EMEA'
COUNTRIES      = []                   # ISO codes to KEEP, e.g. ['BR', 'MX', 'CL'].  [] = all
EXCLUDE_COUNTRIES = []                # ISO codes to DROP, e.g. ['VE', 'AR']
EXCLUDE_SELL   = True                 # drop bonds the house rates Sell
EXCLUDE_SUBORDINATED = True           # senior paper only
ONLY_ATTRACTIVE = False               # True -> only house-view 'attr.' (OP) bonds
MAX_MIN_DENOM  = None                 # e.g. 200_000 -> skip bonds with a bigger minimum piece
MIN_PRICE, MAX_PRICE = None, None     # e.g. 60, 105

# ---- 4. Spread reference — no Treasury feed, nothing to paste -------------
# The reference is the list's OWN curve: the median offer yield of every
# eligible senior, non-Sell bond in each rung window (1y, 2y ... 10y), i.e. a
# proxy for the EMBI Global curve, rebuilt from the feed on every run.
# Because the EMBI Global trades AT LEAST EMBI_OVER_UST_BP above Treasuries,
# a bond sitting on the EM curve is guaranteed at least that much over UST:
#       min spread vs UST  =  (offer yield - EM curve)  +  EMBI_OVER_UST_BP
# So with EMBI_OVER_UST_BP = 100:
#       MIN_SPREAD_BP = 100  -> bonds at or above the EM curve
#       MIN_SPREAD_BP = 150  -> bonds at least 50bp cheap to the EM curve
#       MIN_SPREAD_BP = 50   -> lets in bonds up to 50bp rich to the curve
# Raise EMBI_OVER_UST_BP toward the actual index spread (~200bp) if you want
# the "vs UST" figure to be an estimate rather than a floor.
EMBI_OVER_UST_BP = 100
CURVE_MIN_BONDS  = 5                  # a rung with fewer bonds is interpolated from its neighbours

# ---- 5. Files --------------------------------------------------------------
DATA_DIR   = 'data'
OUTPUT_DIR = 'outputs'
OUTPUT_FILE = None                    # None = auto name, or e.g. 'outputs/my_ladder.xlsx'
MAKE_PDF    = True                    # also write a one-page PDF (same name, .pdf) to email to FAs
PDF_TITLE   = None                    # None = auto, e.g. 'USD Investment Grade EM Bond Ladder, 1-10y'
PDF_LOGO    = DEFAULT_LOGO_PATH       # 'assets/UBS_Logo.png' — skipped silently if absent

# ══════════════════════════════════════════════════════════════════════════════
# End of CONFIG. Everything below just does what the block above says.
# ══════════════════════════════════════════════════════════════════════════════

def _find_optional(patterns_stems, exts, skip=()):
    """Newest file matching any stem in data/, data/current/ or the folder root."""
    cands = []
    for folder in (DATA_DIR, os.path.join(DATA_DIR, 'current'), HERE):
        for stem in patterns_stems:
            for ext in exts:
                cands.extend(glob.glob(os.path.join(folder, f'{stem}.{ext}')))
    cands = [c for c in set(cands)
             if not os.path.basename(c).startswith('~$')
             and not any(s in os.path.basename(c).lower() for s in skip)]
    return max(cands, key=os.path.getmtime) if cands else None


def load_data():
    curr = os.path.join(DATA_DIR, 'current')
    prev = os.path.join(DATA_DIR, 'previous')
    paths = {
        'bond_data':      os.path.join(curr, 'CurrentPublishableBondData.txt'),
        'issuer_data':    os.path.join(curr, 'CurrentPublishableIssuerData.txt'),
        'bond_update':    os.path.join(curr, 'PublishableBondDataUpdate.txt'),
        'issuer_update':  os.path.join(curr, 'PublishableIssuerDataUpdate.txt'),
        'color_flags':    os.path.join(curr, 'PublishableColorFlags.txt'),
        'issuer_texts':   os.path.join(curr, 'IssuerTexts.txt'),
        'issuer_ratings': os.path.join(curr, 'IssuerRatings.txt'),
    }
    missing = [p for p in paths.values() if not os.path.exists(p)]
    if missing:
        sys.exit('Missing feed file(s):\n  ' + '\n  '.join(missing))
    pb, pu = (os.path.join(prev, 'CurrentPublishableBondData.txt'),
              os.path.join(prev, 'PublishableBondDataUpdate.txt'))
    if os.path.exists(pb) and os.path.exists(pu):
        paths['prev_bond_data'], paths['prev_bond_update'] = pb, pu
    paths['priips_ref'] = _find_optional(['*[Pp][Rr][Ii][Ii][Pp][Ss]*'], ('xls', 'xlsx', 'csv'))
    paths['legal_exclusions'] = _find_optional(['*[Ll]egal*', '*[Ee]xclusion*'], ('txt', 'csv'),
                                               skip=('template', 'example', 'sample'))
    return GEMData(paths)


# ---- small helpers -----------------------------------------------------------

def _num(s):
    try:
        return float(str(s).replace(',', ''))
    except (TypeError, ValueError):
        return None


class EMCurve:
    """Median offer yield per rung tenor, linearly interpolated in between,
    flat beyond the ends. Built from the base universe on every run."""

    def __init__(self, rows):
        self.points, self.counts = [], {}
        self.tenors = sorted(set(range(1, 11)) | set(LADDER_YEARS))  # always the full 1-10y grid
        for yr in self.tenors:
            ys = sorted(r['yld'] for r in rows if abs(r['years'] - yr) <= RUNG_WINDOW)
            self.counts[yr] = len(ys)
            if len(ys) >= CURVE_MIN_BONDS:
                mid = len(ys) // 2
                med = ys[mid] if len(ys) % 2 else (ys[mid - 1] + ys[mid]) / 2
                self.points.append((yr, med))
        if not self.points:
            sys.exit('Not enough bonds to build the EM curve — check CURRENCY / data.')

    def at(self, years):
        pts = self.points
        if years <= pts[0][0]:
            return pts[0][1]
        if years >= pts[-1][0]:
            return pts[-1][1]
        for (t0, y0), (t1, y1) in zip(pts, pts[1:]):
            if t0 <= years <= t1:
                return y0 + (y1 - y0) * (years - t0) / (t1 - t0)
        return pts[-1][1]

    def table(self):
        return [(yr, self.at(yr), self.counts.get(yr, 0)) for yr in self.tenors]


def _rating_tier(token):
    return rating_tier(parse_rating(token))


# ---- 1. screen the universe --------------------------------------------------

def build_candidates(data):
    """Two passes over the list.
    Pass 1 — base universe (currency, senior, non-Sell, has maturity/price/
             yield) -> builds the EM curve.  Grade/region/country do NOT
             narrow the curve: it is meant to be the whole EM market.
    Pass 2 — the CONFIG selection filters + spread test on top of that."""
    today = date.today()
    min_rating_tier = _rating_tier(MIN_RATING) if MIN_RATING else None
    if MIN_RATING and min_rating_tier is None:
        sys.exit(f"MIN_RATING '{MIN_RATING}' is not a rating I know (use e.g. 'BB-' or 'Ba3').")

    dropped = {}

    def drop(why):
        dropped[why] = dropped.get(why, 0) + 1

    # ---- pass 1: base universe ------------------------------------------
    base = []
    for b in data.em_bonds:
        row = data.bond_row(b)
        if CURRENCY != 'ANY' and row['ccy'] != CURRENCY.upper():
            drop('currency'); continue
        if EXCLUDE_SUBORDINATED and is_subordinated_bond(b):
            drop('subordinated'); continue
        if EXCLUDE_SELL and row['rec'] == 'SELL':
            drop('Sell'); continue
        if row['maturity_date'] is None:
            drop('perpetual / no maturity'); continue
        years = (row['maturity_date'].date() - today).days / 365.25
        if years <= 0:
            drop('matured'); continue
        px, yld = _num(row['offer_price']), _num(row['offer_yield'])
        if px is None or yld is None:
            drop('no price/yield'); continue
        eff = data.effective_issuer_rating(row['gk'], b)
        row.update(years=years, px=px, yld=yld, _bond=b, _eff=eff,
                   sp=eff['sp_token'] or 'n/a', mdy=eff['mdy_token'] or 'n/a')
        base.append(row)

    curve = EMCurve(base)
    print(f'[ladder] EM curve from {len(base):,} {CURRENCY} bonds '
          f'(median offer yield per rung):')
    print('         ' + '  '.join(f'{yr}y {y:.2f}%({n})' for yr, y, n in curve.table()))

    # ---- pass 2: selection ------------------------------------------------
    out = []
    for row in base:
        itype = row['itype']
        if UNIVERSE == 'sovereign' and itype not in ('SOV', 'SUPRA'):
            drop('universe'); continue
        if UNIVERSE == 'corporate' and itype in ('SOV', 'SUPRA'):
            drop('universe'); continue
        if REGION and row['region'] != REGION:
            drop('region'); continue
        if COUNTRIES and row['country'] not in [c.upper() for c in COUNTRIES]:
            drop('country'); continue
        if row['country'] in [c.upper() for c in EXCLUDE_COUNTRIES]:
            drop('country excluded'); continue
        if ONLY_ATTRACTIVE and row['rec'] != 'OP':
            drop('not attractive'); continue

        is_ig = row['grade'].startswith('Investment')
        if GRADE == 'IG' and not is_ig:
            drop('not IG'); continue
        if GRADE == 'HY' and is_ig:
            drop('not HY'); continue
        if min_rating_tier is not None:
            wt = row['_eff']['worst_tier']
            if wt is None or wt > min_rating_tier:
                drop('below MIN_RATING'); continue

        if MIN_PRICE is not None and row['px'] < MIN_PRICE:
            drop('price < min'); continue
        if MAX_PRICE is not None and row['px'] > MAX_PRICE:
            drop('price > max'); continue
        if MAX_MIN_DENOM is not None:
            md = _num(row['min_denom'].split('/')[0])
            if md is not None and md > MAX_MIN_DENOM:
                drop('min denomination too big'); continue

        ref = curve.at(row['years'])
        vs_curve = (row['yld'] - ref) * 100.0
        ust_floor = vs_curve + EMBI_OVER_UST_BP
        if MIN_SPREAD_BP is not None and ust_floor < MIN_SPREAD_BP:
            drop(f'< {MIN_SPREAD_BP:g}bp over UST'); continue

        row.update(curve=ref, vs_curve_bp=vs_curve, ust_floor_bp=ust_floor)
        out.append(row)

    print(f'[ladder] {len(out):,} bonds pass the filters '
          f'(of {len(data.em_bonds):,} on the list)')
    if dropped:
        print('[ladder] dropped: ' + ', '.join(f'{k} {v}' for k, v in
                                               sorted(dropped.items(), key=lambda kv: -kv[1])))
    return out, curve


# ---- 2. rank within a rung ---------------------------------------------------

def rank_key(rows):
    """Return a sort key for the chosen PICK rule. Lower = better."""
    if PICK == 'yield':
        return lambda r: (-r['yld'], r['px'])
    if PICK == 'price':
        return lambda r: (r['px'], -r['yld'])
    if PICK == 'both':
        by_y = {id(r): i for i, r in enumerate(sorted(rows, key=lambda r: -r['yld']))}
        by_p = {id(r): i for i, r in enumerate(sorted(rows, key=lambda r: r['px']))}
        return lambda r: (by_y[id(r)] + by_p[id(r)], -r['yld'])
    sys.exit(f"PICK must be 'yield', 'price' or 'both' (got {PICK!r})")


# ---- 3. build the ladder -----------------------------------------------------

def build_ladder(cands):
    used_issuers, country_count = set(), {}
    ladder, all_ranked = [], []

    for yr in LADDER_YEARS:
        note = ''
        pool = [r for r in cands if abs(r['years'] - yr) <= RUNG_WINDOW]
        if not pool and RUNG_WINDOW_FALLBACK:
            pool = [r for r in cands if abs(r['years'] - yr) <= RUNG_WINDOW_FALLBACK]
            note = f'widened to +/-{RUNG_WINDOW_FALLBACK}y'
        pool = sorted(pool, key=rank_key(pool))
        for i, r in enumerate(pool, 1):
            all_ranked.append(dict(r, rung=yr, rank=i))

        picked = 0
        for r in pool:
            if ONE_BOND_PER_ISSUER and r['issuer_raw'] in used_issuers:
                continue
            if MAX_PER_COUNTRY and country_count.get(r['country'], 0) >= MAX_PER_COUNTRY:
                continue
            ladder.append(dict(r, rung=yr, note=note))
            used_issuers.add(r['issuer_raw'])
            country_count[r['country']] = country_count.get(r['country'], 0) + 1
            picked += 1
            if picked >= BONDS_PER_RUNG:
                break
        if picked == 0 and pool and ALLOW_REPEAT_IF_NEEDED:
            # Every candidate is an issuer/country already used: better a
            # repeat than a hole in the ladder. Flagged in the Note column.
            r = pool[0]
            ladder.append(dict(r, rung=yr, note=(note + '; ' if note else '')
                               + 'repeat issuer/country (no alternative)'))
            country_count[r['country']] = country_count.get(r['country'], 0) + 1
            picked = 1
        if picked == 0:
            ladder.append({'rung': yr, 'empty': True,
                           'note': 'no eligible bond in window' if not pool
                                   else 'all candidates blocked by issuer/country limits'})
    return ladder, all_ranked


# ---- 4. output ---------------------------------------------------------------

COLS = [  # (header, key, number format)
    ('Rung (y)',        'rung',           '0'),
    ('Issuer',          'issuer',         None),
    ('Country',         'country_display', None),
    ('Type',            'itype',          None),
    ('ISIN / Valor',    'isin_valor',     None),
    ('Ccy',             'ccy',            None),
    ('Coupon',          'coupon',         None),
    ('Maturity',        'maturity',       None),
    ('Yrs to mat.',     'years',          '0.0'),
    ('Offer price',     'px',             '0.0'),
    ('Offer yield %',   'yld',            '0.00'),
    ('EM curve %',      'curve',          '0.00'),
    ('vs EM curve bp',  'vs_curve_bp',    '0'),
    ('Min sprd vs UST bp', 'ust_floor_bp', '0'),
    ("S&P / Moody's",   'ratings',        None),
    ('Grade',           'grade',          None),
    ('View',            'view',           None),
    ('Min denom.',      'min_denom',      None),
    ('Restrictions',    'restrictions',   None),
    ('Note',            'note',           None),
]


def print_ladder(ladder):
    print()
    print(f'  {"Rung":>4}  {"Issuer":<38} {"Maturity":<10} {"Price":>6} {"Yield":>6} '
          f'{"vsEM":>5} {"UST+":>5}  {"Rating":<14} {"View"}')
    for r in ladder:
        if r.get('empty'):
            print(f'  {r["rung"]:>3}y  -- {r["note"]}')
            continue
        print(f'  {r["rung"]:>3}y  {r["issuer"][:38]:<38} {r["maturity"]:<10} '
              f'{r["px"]:>6.1f} {r["yld"]:>6.2f} {r["vs_curve_bp"]:>+5.0f} '
              f'{r["ust_floor_bp"]:>5.0f}  '
              f'{r["ratings"]:<14} {r["view"]}' + (f'   [{r["note"]}]' if r['note'] else ''))
    filled = [r for r in ladder if not r.get('empty')]
    if filled:
        n = len(filled)
        print(f'\n  Equal-weighted: yield {sum(r["yld"] for r in filled)/n:.2f}%  '
              f'price {sum(r["px"] for r in filled)/n:.1f}  '
              f'vs EM curve {sum(r["vs_curve_bp"] for r in filled)/n:+.0f}bp  '
              f'=> at least UST+{sum(r["ust_floor_bp"] for r in filled)/n:.0f}bp  '
              f'({n} of {len(ladder)} rungs filled)')


def write_excel(ladder, ranked, curve, path):
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Alignment
    from openpyxl.utils import get_column_letter

    wb = Workbook()
    hdr_font, hdr_fill = Font(bold=True, color='FFFFFF'), PatternFill('solid', fgColor='4A4A4A')

    def sheet(ws, rows, cols):
        ws.append([h for h, _, _ in cols])
        for c in ws[1]:
            c.font, c.fill, c.alignment = hdr_font, hdr_fill, Alignment(vertical='center', wrap_text=True)
        for r in rows:
            ws.append([r.get(k, '') for _, k, _ in cols])
        for j, (_, k, fmt) in enumerate(cols, 1):
            if fmt:
                for c in ws.iter_rows(min_row=2, min_col=j, max_col=j):
                    c[0].number_format = fmt
            width = max([len(str(r.get(k, ''))) for r in rows] + [len(cols[j-1][0])]) + 2
            ws.column_dimensions[get_column_letter(j)].width = min(max(width, 8), 45)
        ws.freeze_panes = 'B2'

    ws = wb.active
    ws.title = 'Ladder'
    sheet(ws, ladder, COLS)
    filled = [r for r in ladder if not r.get('empty')]
    if filled:
        n = len(filled)
        ws.append([])
        ws.append(['Equal-weighted average', '', '', '', '', '', '', '', '',
                   sum(r['px'] for r in filled) / n,
                   sum(r['yld'] for r in filled) / n,
                   sum(r['curve'] for r in filled) / n,
                   sum(r['vs_curve_bp'] for r in filled) / n,
                   sum(r['ust_floor_bp'] for r in filled) / n])
        for c in ws[ws.max_row]:
            c.font = Font(bold=True)
        for c, fmt in zip(ws[ws.max_row][9:14], ('0.0', '0.00', '0.00', '0', '0')):
            c.number_format = fmt

    sheet(wb.create_sheet('Candidates'), ranked,
          [('Rank', 'rank', '0')] + COLS[:-1])

    ws = wb.create_sheet('Settings')
    ws.append(['Setting', 'Value']); ws['A1'].font = ws['B1'].font = Font(bold=True)
    for k in ('GRADE', 'MIN_SPREAD_BP', 'PICK', 'MIN_RATING', 'LADDER_YEARS', 'RUNG_WINDOW',
              'RUNG_WINDOW_FALLBACK', 'BONDS_PER_RUNG', 'ONE_BOND_PER_ISSUER', 'MAX_PER_COUNTRY', 'ALLOW_REPEAT_IF_NEEDED',
              'CURRENCY', 'UNIVERSE', 'REGION', 'COUNTRIES', 'EXCLUDE_COUNTRIES', 'EXCLUDE_SELL',
              'EXCLUDE_SUBORDINATED', 'ONLY_ATTRACTIVE', 'MAX_MIN_DENOM', 'MIN_PRICE', 'MAX_PRICE',
              'EMBI_OVER_UST_BP', 'CURVE_MIN_BONDS'):
        ws.append([k, str(globals()[k])])
    ws.append([])
    ws.append(['EM curve (rung y)', 'median offer yield %', 'bonds in window'])
    for c in ws[ws.max_row]:
        c.font = Font(bold=True)
    for yr, y, n in curve.table():
        ws.append([yr, round(y, 3), n])
    ws.append(['Min spread vs UST', f'= (yield - EM curve) + {EMBI_OVER_UST_BP}bp '
               f'(EMBI Global trades at least {EMBI_OVER_UST_BP}bp over UST)'])
    ws.append([]); ws.append(['Run date', date.today().isoformat()])
    ws.column_dimensions['A'].width, ws.column_dimensions['B'].width = 24, 60

    os.makedirs(os.path.dirname(path) or '.', exist_ok=True)
    wb.save(path)
    print(f'\n[ladder] written -> {path}')


def _ladder_label():
    grade = {'IG': 'Investment Grade', 'HY': 'High Yield', 'ANY': ''}[GRADE]
    uni = {'sovereign': 'Sovereign', 'corporate': 'Corporate', 'all': ''}[UNIVERSE]
    ccy = '' if CURRENCY == 'ANY' else CURRENCY
    bits = [x for x in (ccy, REGION or '', grade, uni) if x]
    return ' '.join(bits) + f' EM Bond Ladder, {LADDER_YEARS[0]}-{LADDER_YEARS[-1]}y'


def _ladder_rule_text():
    pick = {'yield': 'highest offer yield', 'price': 'lowest offer price',
            'both': 'best combination of yield and price'}[PICK]
    per = 'One bond' if BONDS_PER_RUNG == 1 else f'Up to {BONDS_PER_RUNG} bonds'
    txt = (f'{per} per maturity year from {LADDER_YEARS[0]} to {LADDER_YEARS[-1]} years, '
           f'selected from the CIO Emerging Markets Bond List: {pick} per rung')
    if MIN_SPREAD_BP is not None:
        txt += f', yielding at least {MIN_SPREAD_BP:g}bp over US Treasuries'
    if ONE_BOND_PER_ISSUER:
        txt += ', no issuer repeated'
    if EXCLUDE_SUBORDINATED:
        txt += ', senior bonds only'
    if EXCLUDE_SELL:
        txt += ', no Sell-rated bonds'
    return txt + '.'


def write_pdf(ladder, data, path):
    from reportlab.lib import colors
    from reportlab.lib.pagesizes import A4, landscape
    from reportlab.lib.styles import ParagraphStyle
    from reportlab.lib.units import cm
    from reportlab.lib.enums import TA_RIGHT
    from reportlab.platypus import (SimpleDocTemplate, Paragraph, Spacer, Table,
                                    TableStyle, KeepTogether)

    fonts = register_fonts()
    F, FB, FI = fonts['light'], fonts['bold'], fonts['italic']
    W, H = landscape(A4)
    ML = MR = 1.5 * cm
    title = PDF_TITLE or _ladder_label()
    as_of = data.data_timestamp()

    st_title = ParagraphStyle('t', fontName=FB, fontSize=16, leading=20, textColor=UBS_DARK)
    st_sub   = ParagraphStyle('s', fontName=F,  fontSize=9.5, leading=12, textColor=UBS_MID)
    st_body  = ParagraphStyle('b', fontName=F,  fontSize=8.5, leading=11, textColor=UBS_DARK)
    st_cell  = ParagraphStyle('c', fontName=F,  fontSize=8,   leading=10, textColor=UBS_DARK)
    st_cellr = ParagraphStyle('cr', parent=st_cell, alignment=TA_RIGHT)
    st_hdr   = ParagraphStyle('h', fontName=FB, fontSize=8,   leading=10, textColor=colors.white)
    st_hdrr  = ParagraphStyle('hr', parent=st_hdr, alignment=TA_RIGHT)
    st_note  = ParagraphStyle('n', fontName=F,  fontSize=7,   leading=9, textColor=UBS_MID)
    st_empty = ParagraphStyle('e', fontName=FI, fontSize=8,   leading=10, textColor=UBS_MID)

    def on_page(canv, doc):
        canv.saveState()
        if PDF_LOGO and os.path.exists(PDF_LOGO):
            try:
                canv.drawImage(PDF_LOGO, ML, H - 1.9 * cm, width=2.2 * cm, height=1.1 * cm,
                               preserveAspectRatio=True, mask='auto')
            except Exception:
                pass
        canv.setFont(F, 8); canv.setFillColor(UBS_MID)
        y = H - 1.0 * cm
        for line in (f'Publication date: {date.today().strftime("%d %B %Y")}',
                     'Chief Investment Office GWM', 'Investment Research'):
            canv.drawRightString(W - MR, y, line); y -= 0.35 * cm
        canv.setStrokeColor(UBS_RULE); canv.setLineWidth(0.5)
        canv.line(ML, 1.35 * cm, W - MR, 1.35 * cm)
        canv.setFont(F, 7)
        canv.drawString(ML, 0.9 * cm, 'Source: UBS, rating agencies')
        canv.drawRightString(W - MR, 0.9 * cm, f'Page {doc.page}')
        canv.restoreState()

    # ---- table ---------------------------------------------------------------
    cols = [  # header, key, width(cm), align-right?
        ('Rung',            'rung',          1.2, True),
        ('Issuer',          'issuer',        6.6, False),
        ('ISIN / Valor',    'isin_valor',    3.9, False),
        ('Coupon',          'coupon',        1.6, True),
        ('Maturity',        'maturity',      2.0, True),
        (f'Offer price{INDICATIVE_MARK}', 'px',  1.8, True),
        (f'Offer yield{INDICATIVE_MARK}', 'yld', 1.8, True),
        ("S&amp;P / Moody's", 'ratings',     2.5, False),
        ('CIO view',        'view',          1.5, False),
        ('Min. denom.',     'min_denom',     2.6, True),
        ('Restr.',          'restrictions',  1.2, False),
    ]
    head = [Paragraph(h, st_hdrr if r else st_hdr) for h, _, _, r in cols]
    rows, cmds = [head], []
    for i, r in enumerate(ladder, 1):
        if r.get('empty'):
            cells = [Paragraph(f'{r["rung"]}y', st_cellr),
                     Paragraph('No eligible bond' + (' — all candidates already used'
                               if 'blocked' in r['note'] else ' in this maturity window'), st_empty)
                     ] + [''] * (len(cols) - 2)
            cmds.append(('SPAN', (1, i), (-1, i)))
        else:
            vals = {'rung': f'{r["rung"]}y', 'px': f'{r["px"]:.1f}', 'yld': f'{r["yld"]:.2f}%',
                    'restrictions': r.get('restrictions') or ''}
            cells = [Paragraph(str(vals.get(k, r.get(k, ''))), st_cellr if right else st_cell)
                     for _, k, _, right in cols]
        rows.append(cells)
        if i % 2 == 0:
            cmds.append(('BACKGROUND', (0, i), (-1, i), UBS_LIGHT))

    filled = [r for r in ladder if not r.get('empty')]
    if filled:
        n = len(filled)
        avg = [Paragraph('', st_cell), Paragraph(f'Equal-weighted average ({n} bonds)', ParagraphStyle('a', parent=st_cell, fontName=FB)),
               '', '', '',
               Paragraph(f'{sum(r["px"] for r in filled)/n:.1f}', ParagraphStyle('ar', parent=st_cellr, fontName=FB)),
               Paragraph(f'{sum(r["yld"] for r in filled)/n:.2f}%', ParagraphStyle('ar2', parent=st_cellr, fontName=FB)),
               '', '', '', '']
        rows.append(avg)
        cmds += [('LINEABOVE', (0, len(rows) - 1), (-1, len(rows) - 1), 0.75, UBS_DARK),
                 ('BACKGROUND', (0, len(rows) - 1), (-1, len(rows) - 1), colors.white)]

    tbl = Table(rows, colWidths=[w * cm for _, _, w, _ in cols], repeatRows=1)
    tbl.setStyle(TableStyle([
        ('BACKGROUND',    (0, 0), (-1, 0), UBS_HEADER_BG),
        ('VALIGN',        (0, 0), (-1, -1), 'MIDDLE'),
        ('TOPPADDING',    (0, 0), (-1, -1), 3),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 3),
        ('LEFTPADDING',   (0, 0), (-1, -1), 4),
        ('RIGHTPADDING',  (0, 0), (-1, -1), 4),
        ('LINEBELOW',     (0, 0), (-1, 0), 0.75, UBS_DARK),
        ('LINEBELOW',     (0, 1), (-1, -1), 0.25, UBS_RULE),
    ] + cmds))

    story = [
        Spacer(1, 0.6 * cm),
        Paragraph(title, st_title),
        Paragraph('Chief Investment Office GWM  |  Emerging Markets Bond List', st_sub),
        Spacer(1, 0.25 * cm),
        Paragraph(_ladder_rule_text(), st_body),
        Spacer(1, 0.35 * cm),
        tbl,
        Spacer(1, 0.3 * cm),
        Paragraph(f'{INDICATIVE_MARK} Indicative values. Market data as of {as_of}; prices and '
                  f'yields are indicative only and may not be available in every jurisdiction. '
                  f'Ratings shown as S&amp;P / Moody\'s. CIO view: attr. = attractive, fair = fair value, '
                  f'exp. = expensive. Restr.: 1 = MiFID complex product, 2 = PRIIPs-relevant, no KID. '
                  f'Minimum denomination shown as minimum amount / increment.', st_note),
    ]
    if MIN_SPREAD_BP is not None:
        story.append(Paragraph(
            f'Spread screen: yield compared with the median yield of the EM Bond List at the same '
            f'maturity, assuming the EMBI Global trades at least {EMBI_OVER_UST_BP}bp over US '
            f'Treasuries; bonds shown are therefore at least {MIN_SPREAD_BP:g}bp over Treasuries.',
            st_note))

    os.makedirs(os.path.dirname(path) or '.', exist_ok=True)
    doc = SimpleDocTemplate(path, pagesize=(W, H), leftMargin=ML, rightMargin=MR,
                            topMargin=1.6 * cm, bottomMargin=1.7 * cm,
                            title=title, author='UBS CIO GWM')
    doc.build(story, onFirstPage=on_page, onLaterPages=on_page)
    print(f'[ladder] written -> {path}')


# ---- 5. CLI overrides (optional; CONFIG is the default) ----------------------

def apply_cli():
    ap = argparse.ArgumentParser(description='Build a 1y-10y EM bond ladder from the EM Bond List.')
    ap.add_argument('--grade', choices=['IG', 'HY', 'ANY'])
    ap.add_argument('--spread', type=float, metavar='BP', help='min spread over UST, e.g. 100')
    ap.add_argument('--pick', choices=['yield', 'price', 'both'])
    ap.add_argument('--min-rating', metavar='RATING')
    ap.add_argument('--ccy'); ap.add_argument('--universe', choices=['sovereign', 'corporate', 'all'])
    ap.add_argument('--region'); ap.add_argument('--country', action='append', metavar='ISO')
    ap.add_argument('--per-rung', type=int); ap.add_argument('--years', help='e.g. 1-10 or 2-7')
    ap.add_argument('--out')
    a = ap.parse_args()
    g = globals()
    if a.grade:      g['GRADE'] = a.grade
    if a.spread is not None: g['MIN_SPREAD_BP'] = a.spread
    if a.pick:       g['PICK'] = a.pick
    if a.min_rating: g['MIN_RATING'] = a.min_rating
    if a.ccy:        g['CURRENCY'] = a.ccy.upper()
    if a.universe:   g['UNIVERSE'] = a.universe
    if a.region:     g['REGION'] = a.region
    if a.country:    g['COUNTRIES'] = [c.upper() for c in a.country]
    if a.per_rung:   g['BONDS_PER_RUNG'] = a.per_rung
    if a.years:
        lo, hi = (int(x) for x in a.years.split('-'))
        g['LADDER_YEARS'] = list(range(lo, hi + 1))
    if a.out:        g['OUTPUT_FILE'] = a.out


def main():
    apply_cli()
    print(f'[ladder] GRADE={GRADE}  PICK={PICK}  MIN_SPREAD_BP={MIN_SPREAD_BP}  '
          f'CCY={CURRENCY}  UNIVERSE={UNIVERSE}  REGION={REGION or "all"}  '
          f'rungs={LADDER_YEARS[0]}-{LADDER_YEARS[-1]}y')
    if MIN_SPREAD_BP is not None:
        print(f'[ladder] spread test: (yield - EM curve) + {EMBI_OVER_UST_BP}bp '
              f'>= {MIN_SPREAD_BP:g}bp  i.e. bonds at least '
              f'{MIN_SPREAD_BP - EMBI_OVER_UST_BP:+g}bp vs the EM curve')
    data = load_data()
    cands, curve = build_candidates(data)
    ladder, ranked = build_ladder(cands)
    print_ladder(ladder)
    out = OUTPUT_FILE or os.path.join(
        OUTPUT_DIR, f'Bond_Ladder_{GRADE}_{PICK}'
        + (f'_{MIN_SPREAD_BP:g}bp' if MIN_SPREAD_BP is not None else '')
        + (f'_{REGION.replace(" ", "")}' if REGION else '') + '.xlsx')
    write_excel(ladder, ranked, curve, out)
    if MAKE_PDF:
        write_pdf(ladder, data, os.path.splitext(out)[0] + '.pdf')
    return ladder


if __name__ == '__main__':
    main()
