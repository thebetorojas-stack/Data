#!/usr/bin/env python3
"""
Emerging Markets Bond List — PDF Generator v3
==============================================

Cleaner rewrite of gem_report_builder_v2.py with:

    • Top-of-file CONFIG block for every business constant (currency, country,
      green-bond, analyst, region, legal text, fonts, logo)
    • Running header (title / subtitle / analyst block) on every page via
      PageTemplate callbacks — no ad-hoc rendering inside each section
    • Clickable table of contents on the cover page (hyperlinks jump to sections)
    • Proper issuer-name resolution (CTL_List_Name, then CML_Publish_Name)
    • Country-code → English-name map; sovereigns rendered without parens
    • Full currency names for reference-list section headers
    • Issuer-type sort rank (SOV/SUPRA before FIN before CORP) within each group
    • Investment-grade / Speculative-grade split with risk-warning legend
    • Reference lists grouped by Currency × Region × Grade
    • Page numbers (Page X of Y) and centered source line on every page
    • "¹ Indicative values" footnote under each bond-list table
    • Frutiger 45 Light registered when available, Helvetica fallback otherwise

Usage
-----
    python3 gem_report_builder_v3.py \\
        --bond-data        CurrentPublishableBondData.txt \\
        --issuer-data      CurrentPublishableIssuerData.txt \\
        --bond-update      PublishableBondDataUpdate.txt \\
        --issuer-update    PublishableIssuerDataUpdate.txt \\
        --color-flags      PublishableColorFlags.txt \\
        --issuer-texts     IssuerTexts.txt \\
        --issuer-ratings   IssuerRatings.txt \\
        --prev-bond-data   CurrentPublishableBondData-a7203bca.txt \\
        --prev-bond-update PublishableBondDataUpdate-bf12ddf2.txt \\
        --logo             assets/UBS_Logo.png \\
        --output           GEM_List.pdf


MAINTAINER'S GUIDE — where to change common things
---------------------------------------------------
This module is organised into the numbered sections below (search for the
"# N." banners). For the most frequent edits:

  • Add/rename a COUNTRY name .............. edit COUNTRY_NAMES        (section 2)
  • Re-classify an issuer's REGION ......... edit REGION_NAME_OVERRIDES(section 5)
      (name-based override that wins over the country-code REGION_MAP)
  • Add/rename a CURRENCY label ............ edit CURRENCY_NAMES       (section 2)
  • Hide a currency from reference lists ... edit EXCLUDED_CURRENCIES  (section 2)
  • Update the analyst roster (cover page) . edit ANALYST_ROSTER       (section 2)
  • Change what counts as "subordinated" ... edit SUBORDINATED_*       (section 2)
  • Change IG vs HY rating cutoff .......... edit RATING_SCALE / IG_MAX_TIER (sec 2)
  • Change which bonds make the list ....... edit GEMData._classify_for_list (sec 6)
  • Change which bonds are "Top List" ...... edit GEMData.top_list_bonds    (sec 6)
  • Change which bonds hit reference lists . edit GEMData.reference_list_bonds(sec 6)
  • Change a single bond row's displayed
    values (ratings, region, yield, etc.) .. edit GEMData.bond_row    (section 6)

DATA SOURCES (how the seven input files relate)
  • bond-data      — one row per security (price, coupon, maturity, ISIN…)
  • bond-update    — weekly per-security overlay: recommendation, WMRFlag,
                     and the BOND-LEVEL agency ratings (RatingSP/RatingMdy)
                     used for subordinated bonds
  • issuer-data    — one row per issuer (name, country, type, issuer ratings)
  • issuer-update  — weekly per-issuer overlay
  • color-flags    — per-issuer credit-view colour grid (issuer-descriptions page)
  • issuer-texts   — issuer description paragraphs
  • issuer-ratings — agency issuer ratings (S&P / Moody's)
  • prev-bond-*    — last week's bond files, used only for the week-on-week
                     Changes / Additions / Deletions diff
"""

# ══════════════════════════════════════════════════════════════════════════════
# 1. IMPORTS
# ══════════════════════════════════════════════════════════════════════════════

import argparse
import csv
import os
from collections import Counter, defaultdict
from datetime import datetime, date

from reportlab.lib import colors
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import cm, mm
from reportlab.lib.enums import TA_LEFT, TA_CENTER, TA_RIGHT
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.platypus import (
    BaseDocTemplate, Flowable, Frame,
    NextPageTemplate, PageBreak, PageTemplate, Paragraph, Spacer,
    Table, TableStyle,
)
from reportlab.platypus.doctemplate import ActionFlowable


# ══════════════════════════════════════════════════════════════════════════════
# 2. USER-EDITABLE CONFIGURATION
#    Everything in this block is business-data that changes without code logic.
#    Edit here rather than in render code.
# ══════════════════════════════════════════════════════════════════════════════

# ---- Currency display names (section headers for Reference lists) ------------
# Codes that already read naturally (USD/EUR/GBP/CHF/AUD/HKD) stay as codes.
# Exotics expand to their English names.
CURRENCY_NAMES = {
    'USD': 'USD',
    'EUR': 'EUR',
    'GBP': 'GBP',
    'CHF': 'CHF',
    'AUD': 'AUD',
    'HKD': 'HKD',
    'CAD': 'CAD',
    'CNY': 'Chinese renminbi',
    'SGD': 'Singapore dollar',
    'MXN': 'Mexican peso',
    'BRL': 'Brazilian real',
    'ZAR': 'South African rand',
    'IDR': 'Indonesian rupiah',
    'INR': 'Indian rupee',
    'TRY': 'Turkish lira',
    'PLN': 'Polish zloty',
    'JPY': 'Japanese yen',
    'NOK': 'Norwegian krone',
    'SEK': 'Swedish krona',
    'DKK': 'Danish krone',
    'ARS': 'Argentine peso',
    'COP': 'Colombian peso',
    'MXV': 'Mexican UDI',
    'RUB': 'Russian rouble',
    'PEN': 'Peruvian sol',
    'CLP': 'Chilean peso',
    'HUF': 'Hungarian forint',
    'CZK': 'Czech koruna',
    'THB': 'Thai baht',
    'MYR': 'Malaysian ringgit',
    'PHP': 'Philippine peso',
    'KRW': 'Korean won',
    'TWD': 'Taiwan dollar',
    'AED': 'UAE dirham',
    'SAR': 'Saudi riyal',
    'ILS': 'Israeli shekel',
    'NZD': 'New Zealand dollar',
    'RON': 'Romanian leu',
}

# ---- Country code → English name (used everywhere countries appear) ----------
COUNTRY_NAMES = {
    'AE': 'United Arab Emirates',
    'AN': 'Netherlands Antilles',        # obsolete ISO code, but still present in source data
    'AR': 'Argentina', 'AT': 'Austria', 'AU': 'Australia',
    'BE': 'Belgium', 'BH': 'Bahrain', 'BM': 'Bermuda', 'BR': 'Brazil',
    'BS': 'Bahamas', 'CA': 'Canada', 'CH': 'Switzerland',
    'CI': "Côte d'Ivoire",
    'CL': 'Chile',
    'CN': 'China', 'CO': 'Colombia', 'CR': 'Costa Rica', 'CY': 'Cyprus',
    'CZ': 'Czech Republic', 'DE': 'Germany', 'DK': 'Denmark',
    'DO': 'Dominican Republic', 'EC': 'Ecuador', 'EG': 'Egypt',
    'ES': 'Spain', 'FI': 'Finland', 'FR': 'France', 'GB': 'United Kingdom',
    'GR': 'Greece', 'HK': 'Hong Kong',
    'HN': 'Honduras',
    'HR': 'Croatia',
    'HU': 'Hungary', 'ID': 'Indonesia',
    'IE': 'Ireland', 'IL': 'Israel', 'IM': 'Isle of Man', 'IN': 'India', 'IT': 'Italy',
    'JM': 'Jamaica', 'JP': 'Japan',
    'KE': 'Kenya',
    'KR': 'Korea', 'KW': 'Kuwait',
    'KY': 'Cayman Islands', 'KZ': 'Kazakhstan', 'LK': 'Sri Lanka',
    'LU': 'Luxembourg', 'MA': "Morocco", 'MO': 'Macao', 'MX': 'Mexico', 'MY': 'Malaysia',
    'NG': 'Nigeria', 'NL': 'Netherlands', 'NO': 'Norway', 'NZ': 'New Zealand',
    'OM': 'Oman', 'PA': 'Panama', 'PE': 'Peru', 'PH': 'Philippines',
    'PK': 'Pakistan', 'PL': 'Poland', 'PT': 'Portugal', 'PY': 'Paraguay',
    'QA': 'Qatar', 'RO': 'Romania', 'RS': 'Serbia', 'RU': 'Russia',
    'SA': 'Saudi Arabia', 'SE': 'Sweden', 'SG': 'Singapore',
    'TH': 'Thailand', 'TR': 'Türkiye', 'TT': 'Trinidad and Tobago',
    'TW': 'Taiwan',
    'TZ': 'Tanzania',
    'UA': 'Ukraine',
    'US': 'United States',
    'UY': 'Uruguay',
    'UZ': 'Uzbekistan', 'VE': 'Venezuela', 'VN': 'Vietnam', 'ZA': 'South Africa',
}

# ---- MiFID / PRIIPs restriction rules ---------------------------------------
# The "Restrictions" column on each bond row shows a comma-separated list of
# numeric codes. The codes map to the legend printed on the guidance page:
#   1) Complex bond under MiFID
#   2) PRIIPS relevant bond, KID missing
#
# No explicit flag exists in the raw data files, so we derive both codes from
# the bond characteristics.
#
# Rule for "1" — MiFID "Complex" bond. Applied when the bond has any embedded
# derivative or non-vanilla feature:
#   • `redeemable = Y`  (the issuer can call the bond — callable)
#   • `retractable = Y` (the holder can put the bond — putable)
#   • Cover type ∈ complex set (subordinated / hybrid / CoCo / etc.)
#   • FOType contains "callable", "convertible", "perpetual", "structured"
#
# A plain vanilla senior bullet bond (redeemable=N, retractable=N, SEN cover
# type, FOType="straight bond") is NOT complex — its restrictions cell stays
# empty even if the minimum denomination is high.
MIFID_COMPLEX_COVER_TYPES = {'SUB', 'PER', 'HYP', 'CCN'}
MIFID_COMPLEX_FOTYPE_KEYWORDS = ('callable', 'convertible', 'perpetual',
                                 'structured', 'subordinated', 'hybrid')

# ---- Subordinated-bond identification ---------------------------------------
# Covered-type codes that count as "subordinated" for the exclusion rule
# documented on the guidance page. Hybrid (HYP) is included because hybrid
# capital instruments are by construction subordinated to senior debt.
SUBORDINATED_COVER_TYPES = {'SUB', 'PER', 'CCN', 'HYP'}
SUBORDINATED_FOTYPE_KEYWORDS = ('subordinated', 'hybrid')

# ---- Rating normalization (S&P/Fitch ↔ Moody's, common numeric tier) -------
# Lower tier = better credit. IG cutoff: tier <= IG_MAX_TIER.
RATING_SCALE = {
    # S&P / Fitch — Investment Grade
    'AAA': 0,
    'AA+': 1, 'AA': 2, 'AA-': 3,
    'A+':  4, 'A':  5, 'A-':  6,
    'BBB+':7, 'BBB':8, 'BBB-':9,
    # S&P / Fitch — Speculative Grade
    'BB+':10, 'BB':11, 'BB-':12,
    'B+': 13, 'B': 14, 'B-': 15,
    'CCC+':16,'CCC':17,'CCC-':18,
    'CC': 19, 'C':  20, 'D':  21,
    # Moody's — Investment Grade
    'Aaa': 0,
    'Aa1': 1, 'Aa2': 2, 'Aa3': 3,
    'A1':  4, 'A2':  5, 'A3':  6,
    'Baa1':7, 'Baa2':8, 'Baa3':9,
    # Moody's — Speculative Grade
    'Ba1':10, 'Ba2':11, 'Ba3':12,
    'B1': 13, 'B2': 14, 'B3': 15,
    'Caa1':16,'Caa2':17,'Caa3':18,
    'Ca': 19,
}
IG_MAX_TIER = 9   # BBB- / Baa3 = lowest investment-grade tier
# Tokens that explicitly mean "no rating available" — not the same as missing.
NO_RATING_TOKENS = {'N.A.', 'NA', 'N/A', 'NR', 'WR', '-', ''}

# Rule for "2" — PRIIPs KID missing. For EM bonds the issuer is almost always
# outside the EEA and therefore does not register a PRIIPs KID. As a result,
# for any PRIIPs-relevant bond (which in the EM universe effectively means any
# bond that also trips Rule 1), restriction "2" co-applies. Set this to False
# if you have a reliable KID-availability source and want restriction "2" to
# apply ONLY to ISINs in KID_MISSING_ISINS.
PRIIPS_KID_AUTO_APPLY = True
KID_MISSING_ISINS = set()   # used only when PRIIPS_KID_AUTO_APPLY = False

# ---- Currencies to EXCLUDE from reference lists ------------------------------
# Reference-list pages (Bonds in X, Region) are NOT generated for these codes.
# Add or remove codes here to change the report without touching render code.
EXCLUDED_CURRENCIES = {
    'ZAR',  # South African rand
    'IDR',  # Indonesian rupiah
    'INR',  # Indian rupee
    'TRY',  # Turkish lira
    'NOK',  # Norwegian krone
    'ARS',  # Argentine peso
    'COP',  # Colombian peso
    'MXV',  # Mexican UDI
    'RUB',  # Russian rouble
}

# ---- EM region map (bonds are grouped Asia / EMEA / GCC / Latin America) -----
REGION_MAP = {
   'Asia': {'CN', 'HK', 'IN', 'ID', 'KR', 'MY', 'PH', 'SG', 'TH', 'TW',
            'VN', 'PK', 'BD', 'LK', 'MN', 'MO', 'JP', 'AU', 'NZ'},
   'Latin America': {'BR', 'MX', 'CL', 'CO', 'PE', 'AR', 'PA', 'UY', 'CR',
                     'DO', 'GT', 'EC', 'JM', 'TT', 'VE', 'PY', 'BB', 'BM',
                     'KY', 'BS',
                     # Central America / Caribbean additions
                     'HN',  # Honduras — fixes CABEI
                     'NI',  # Nicaragua
                     'SV',  # El Salvador
                     'BO',  # Bolivia
                     'GY',  # Guyana
                     'SR',  # Suriname
                     'BZ',  # Belize
                     'AW',  # Aruba
                     'CW',  # Curaçao
                     'HT',  # Haiti
                     },
   'GCC': {'AE', 'SA', 'QA', 'KW', 'BH', 'OM'},
   # Note: 'AN' (Netherlands Antilles) used to live here. Removed because the
   # code is obsolete (ISO retired AN in 2010) and several issuers tagged 'AN'
   # are actually Dutch (Netherlands proper) — e.g. Prosus N.V.
   # Everything else in the EMMA universe falls through to 'EMEA'.
}

# ---- Green / Social / Sustainability label map (bond-data column GreenBond) --
GREEN_LABELS = {
    'G': 'Green',
    'S': 'Social',
    'U': 'Sustainability',
}

# ---- Issuer-type sort rank (within each region × grade block) ----------------
ISSUER_TYPE_RANK = {'SOV': 0, 'SUPRA': 0, 'FIN': 1, 'CORP': 2}

# ---- Acronyms / lower-case connectives for Title-Case filter -----------------
TITLE_ACRONYMS = {'EU', 'UK', 'US', 'USA', 'UAE', 'CAF', 'EMEA', 'APAC',
                  'GCC', 'ASEAN', 'DR', 'OCBC', 'AIA', 'KBC', 'BNP',
                  'UBS', 'HSBC', 'ICBC', 'NATO'}
TITLE_LOWERS = {'of', 'the', 'and', 'for', 'in', 'on', 'a', 'an', 'de',
                'la', 'del', 'y'}

# ---- Analyst display-name overrides (optional) -------------------------------
# Key = value as it appears in PublishableBondDataUpdate.BondAnalyst
# (usually just a surname). Override only when the data has a typo/abbrev.
ANALYST_DISPLAY_NAMES = {
    # 'McLauchlan': 'McLauchlan',
}

# ---- Analyst roster for cover page ("Analysts' area of expertise") -----------
# This is the authoritative list shown on page 1. Edit here when analysts join
# or leave the team. Each entry is (Display name, Title / branch, Expertise).
# The order on the page follows this list order.
ANALYST_ROSTER = [
    ('Alberto Rojas',           'Investment Strategist, CIO Americas, UBS Financial Services Inc. (UBS FS)',    'Sovereign bonds in Latin America, Bond list manager'),
    ('Alejo Czerwonko',         'Chief Investment Officer EM Americas, UBS Financial Services Inc. (UBS FS)',  'Sovereign bonds in Emerging Markets'),
    ('Clarissa Chow',           'CFA, Analyst, UBS AG Singapore Branch',                                       'South Asia corporate bonds'),
    ('Devinda Paranathanthri',  'Credit Strategist, UBS AG Singapore Branch',                                  'South Asia corporate and sovereign bonds'),
    ('Donald McLauchlan',       'LatAm Credit Strategist, UBS Financial Services Inc. (UBS FS)',               'Corporate bonds in Latin America'),
    ('Emre Tekmen',             'CIO Credit Analyst CEEMEA, UBS Switzerland AG',                               'EMEA sovereign bonds'),
    ('Eve Li',                  'Credit Strategist, UBS AG Hong Kong Branch',                                  'North Asia corporate bonds'),
    ('Joel Tan',                'Credit Strategist, UBS AG Singapore Branch',                                  'South Asia corporate bonds'),
    ('Laura Assis Iragorri',    'Credit Strategist, UBS Mexico City Branch',                                   'Sovereign bonds in Latin America'),
    ('Santosh Bukitgar',        'CIO Credit Analyst CEEMEA, UBS Singapore Branch',                             'CEEMEA corporate bonds'),
    ('Tatiana Boroditskaya',    'PhD, Analyst, UBS AG London Branch',                                          'EMEA corporate bonds'),
    ('Timothy Tay',             'Chief Investment Officer Credit APAC, UBS AG Singapore Branch',               'Asia corporate bonds'),
    ('Zixuan Liu',              'Credit Strategist, UBS AG Singapore Branch',                                  'North Asia corporate bonds'),
]

# ---- Market-data note on cover page ------------------------------------------
# The "{timestamp}" placeholder is filled at build time with the current moment
# (when the PDF is generated).
MARKET_DATA_NOTE = (
    'Note that the bonds included in this publication may not necessarily be registered '
    'or available in your specific jurisdiction. Market data shown in this publication '
    'is as of {timestamp}. Please note that prices, yields etc. are indicative values only.'
)

# Timezone label appended to the creation timestamp on the cover page.
# Change if the build host runs in a different zone.
PUBLICATION_TZ_LABEL = 'CET'

# ---- Fonts ------------------------------------------------------------------
FRUTIGER_FILES = {
    'light':       'Frutiger45Light.ttf',
    'light_italic':'Frutiger45LightItalic.ttf',
    'bold':        'Frutiger65Bold.ttf',
    'bold_italic': 'Frutiger65BoldItalic.ttf',
}
FRUTIGER_DIR = 'fonts'

# ---- Colours (UBS CIO palette) ----------------------------------------------
UBS_DARK       = colors.HexColor('#262626')
UBS_MID        = colors.HexColor('#4A4A4A')
UBS_LIGHT      = colors.HexColor('#F2F2F2')
UBS_HEADER_BG  = colors.HexColor('#6B7B6B')   # olive-gray from published PDF
UBS_SECTION_BG = colors.HexColor('#E0E0E0')   # issuer-group row
UBS_SUBGROUP_BG= colors.HexColor('#EDEDED')   # sub-group
UBS_RULE       = colors.HexColor('#AFAFAF')
UBS_RED        = colors.HexColor('#B22222')
UBS_GREEN      = colors.HexColor('#2E7D32')
UBS_YELLOW     = colors.HexColor('#F9A825')

# ---- Page geometry ----------------------------------------------------------
PAGE_WIDTH, PAGE_HEIGHT = landscape(A4)
MARGIN_LEFT   = 1.0 * cm
MARGIN_RIGHT  = 1.0 * cm
MARGIN_TOP    = 2.2 * cm          # room for running header
MARGIN_BOTTOM = 1.3 * cm          # room for running footer
HEADER_Y      = PAGE_HEIGHT - 1.0 * cm
FOOTER_Y      = 0.7 * cm
CONTENT_WIDTH = PAGE_WIDTH - MARGIN_LEFT - MARGIN_RIGHT
CONTENT_HEIGHT= PAGE_HEIGHT - MARGIN_TOP - MARGIN_BOTTOM

# ---- Assets / strings -------------------------------------------------------
DEFAULT_LOGO_PATH = 'assets/UBS_Logo.png'

REPORT_TITLE    = 'Emerging Markets Bond List'
REPORT_SUBTITLE = 'Chief Investment Office GWM'

SOURCE_DEFAULT  = 'Source: UBS'
SOURCE_RATINGS  = 'Source: UBS, rating agencies'

DISCLAIMER_PAGE_1 = (
    'This report has been prepared by UBS Switzerland AG, UBS Financial Services Inc. (UBS FS). '
    'Analyst certification and required disclosures begin on page {p}.'
)

SPECULATIVE_LEGEND = (
    'These issuers are more risky. Their ability to meet payments in the future is '
    'questionable, see rating definitions for details.'
)

TOP_LIST_INTRO = (
    'The CIO Top Emerging Markets Bond List provides guidance on our highest conviction '
    'EM bonds under coverage. Our approach combines bottom-up insights on issuers and bonds '
    'with tactical top-down calls.'
)

# Footnote symbol for Offer price / Offer yield columns. An asterisk is used
# (rather than the superscript 1) so there is no visual collision with the
# digit "1" that can appear in the Restrictions column.
INDICATIVE_MARK = '*'
INDICATIVE_FOOTNOTE = f'{INDICATIVE_MARK} Indicative values'


# ══════════════════════════════════════════════════════════════════════════════
# 3. FONT REGISTRATION (Frutiger with Helvetica fallback)
# ══════════════════════════════════════════════════════════════════════════════

DEJAVU_CANDIDATES = [
    '/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf',
    '/usr/share/fonts/dejavu/DejaVuSans.ttf',
    '/Library/Fonts/DejaVuSans.ttf',
    '/System/Library/Fonts/Supplemental/DejaVuSans.ttf',
]


def _register_symbol_font():
    """Register a Unicode-rich font for symbol glyphs (arrows, etc.) the
    main text font may not cover. Returns the registered name or None.
    Tries DejaVuSans first, then falls back to None — callers should detect
    the None and use ASCII alternatives.
    """
    for path in DEJAVU_CANDIDATES:
        if os.path.exists(path):
            try:
                pdfmetrics.registerFont(TTFont('SymbolText', path))
                return 'SymbolText'
            except Exception:
                pass
    return None


def register_fonts(frutiger_dir: str = FRUTIGER_DIR):
    """Register Frutiger family if TTFs are present; otherwise fall back to Helvetica.

    Returns a dict of logical names → registered font names so the rest of the
    code can use BASE_FONT['light'] etc. without caring which family won.
    Also includes a 'symbol' entry for a Unicode-rich fallback font (e.g.
    arrow glyphs). 'symbol' may be None if no suitable font is found.
    """
    symbol = _register_symbol_font()
    try:
        for key, filename in FRUTIGER_FILES.items():
            path = os.path.join(frutiger_dir, filename)
            if not os.path.exists(path):
                raise FileNotFoundError(path)
        pdfmetrics.registerFont(TTFont('Frutiger-Light',    os.path.join(frutiger_dir, FRUTIGER_FILES['light'])))
        pdfmetrics.registerFont(TTFont('Frutiger-LightIt',  os.path.join(frutiger_dir, FRUTIGER_FILES['light_italic'])))
        pdfmetrics.registerFont(TTFont('Frutiger-Bold',     os.path.join(frutiger_dir, FRUTIGER_FILES['bold'])))
        pdfmetrics.registerFont(TTFont('Frutiger-BoldIt',   os.path.join(frutiger_dir, FRUTIGER_FILES['bold_italic'])))
        pdfmetrics.registerFontFamily(
            'Frutiger',
            normal='Frutiger-Light', bold='Frutiger-Bold',
            italic='Frutiger-LightIt', boldItalic='Frutiger-BoldIt',
        )
        print('[fonts] Frutiger 45 Light registered')
        return {
            'light': 'Frutiger-Light', 'bold': 'Frutiger-Bold',
            'italic': 'Frutiger-LightIt', 'bold_italic': 'Frutiger-BoldIt',
            'symbol': symbol,
        }
    except (FileNotFoundError, Exception) as e:
        print(f'[fonts] Frutiger unavailable ({e}); falling back to Helvetica')
        return {
            'light': 'Helvetica', 'bold': 'Helvetica-Bold',
            'italic': 'Helvetica-Oblique', 'bold_italic': 'Helvetica-BoldOblique',
            'symbol': symbol,
        }


# ══════════════════════════════════════════════════════════════════════════════
# 4. DATA LOADING
# ══════════════════════════════════════════════════════════════════════════════

def _open(path):
    return open(path, 'r', encoding='utf-8-sig', errors='replace')

def _rows_csv(path, **kw):
    with _open(path) as f:
        yield from csv.DictReader(f, **kw)

def load_bonds(path):
    """CurrentPublishableBondData.txt — comma-delimited, quoted."""
    return [dict(r) for r in _rows_csv(path, quotechar='"')]

def load_issuers(path):
    """CurrentPublishableIssuerData.txt — comma-delimited, quoted.
    Returns {gk: row}."""
    out = {}
    for r in _rows_csv(path, quotechar='"'):
        gk = (r.get('GK_Nummer') or '').strip()
        if gk:
            out[gk] = r
    return out

def load_bond_updates(path):
    """PublishableBondDataUpdate.txt — semicolon-delimited.
    Returns {isin: row}."""
    out = {}
    for r in _rows_csv(path, delimiter=';'):
        isin = (r.get('ISIN') or '').strip()
        if isin:
            out[isin] = r
    return out

def load_issuer_updates(path):
    """PublishableIssuerDataUpdate.txt — semicolon-delimited.
    Returns {gk: row}."""
    out = {}
    for r in _rows_csv(path, delimiter=';'):
        gk = (r.get('GK_Nummer') or '').strip()
        if gk:
            out[gk] = r
    return out

def load_color_flags(path):
    """PublishableColorFlags.txt — semicolon-delimited.
    Returns {gk: {covertype: row}}."""
    out = defaultdict(dict)
    for r in _rows_csv(path, delimiter=';'):
        gk = (r.get('GK') or '').strip()
        ct = (r.get('CoverType') or '').strip()
        if gk:
            out[gk][ct] = r
    return out

def load_issuer_texts(path):
    """IssuerTexts.txt — semicolon-delimited.
    Returns {gk: [rows...]}."""
    out = defaultdict(list)
    for r in _rows_csv(path, delimiter=';'):
        gk = (r.get('GKNo') or '').strip()
        if gk:
            out[gk].append(r)
    return out

def load_issuer_ratings(path):
    """IssuerRatings.txt — semicolon-delimited (Expr1000;MDY;SP;Block).
    Returns {gk: row}."""
    out = {}
    for r in _rows_csv(path, delimiter=';'):
        gk = (r.get('Expr1000') or '').strip()
        if gk:
            out[gk] = r
    return out


# ══════════════════════════════════════════════════════════════════════════════
# 5. NAME / LABEL HELPERS
# ══════════════════════════════════════════════════════════════════════════════

def title_case_name(s: str) -> str:
    """Convert ALL-CAPS or mixed-case name to Title Case, preserving acronyms."""
    if not s:
        return s
    words = s.split()
    out = []
    for i, w in enumerate(words):
        u = w.upper().strip('.,()')
        if u in TITLE_ACRONYMS:
            out.append(u)
        elif i > 0 and w.lower() in TITLE_LOWERS:
            out.append(w.lower())
        else:
            out.append(w.capitalize())
    return ' '.join(out)


def country_name(code: str) -> str:
    c = (code or '').strip().upper()
    return COUNTRY_NAMES.get(c, c)


def currency_name(ccy: str) -> str:
    c = (ccy or '').strip().upper()
    return CURRENCY_NAMES.get(c, c)


def analyst_name(raw: str) -> str:
    if not raw:
        return ''
    raw = raw.strip()
    return ANALYST_DISPLAY_NAMES.get(raw, raw)


def em_region(code: str) -> str:
    c = (code or '').strip().upper()
    for region, members in REGION_MAP.items():
        if c in members:
            return region
    return 'EMEA'


# ---- Issuer-name → region overrides ----------------------------------------
# Some issuers belong to a different analyst-coverage region than their legal
# domicile suggests. The canonical examples:
#   • Hyundai Capital America (US-domiciled but covered by LatAm credit team)
#   • Newmont (US-domiciled gold miner with predominantly Latin American mines)
#   • Vedanta Resources (India-domiciled, sometimes incorporated elsewhere)
#   • Genting Overseas (Malaysian conglomerate's offshore vehicle)
#   • Mercado Libre (Argentine/Uruguayan tech, sometimes coded US)
# Add to either tuple to re-classify by issuer-name match (case-sensitive
# substring search against the issuer display name).
REGION_NAME_OVERRIDES = {
    'Asia':          ('Vedanta Res', 'Genting Overseas'),
    'Latin America': ('Hyundai Capital America', 'Newmont', 'Mercado Libre'),
}


def effective_region(name: str, cc: str) -> str:
    """Resolve the region for an issuer using name overrides first, then the
    country-code map. Used by both the PDF and Excel builders so the two
    outputs always agree on the regional placement of every bond."""
    if name:
        for region, patterns in REGION_NAME_OVERRIDES.items():
            if any(p in name for p in patterns):
                return region
    return em_region(cc)


def format_percent(raw) -> str:
    try:
        v = float(raw)
        return f'{round(v,3):g}%'
    except (TypeError, ValueError):
        return (raw or '').strip()


def format_price(raw) -> str:
    try:
        v = float(raw)
        return f'{v:.1f}'
    except (TypeError, ValueError):
        return (raw or '').strip()


# ─── Rating helpers ──────────────────────────────────────────────────────────

def parse_rating(raw):
    """Normalize a raw rating string to its base symbol (e.g. 'BBB+', 'Baa3').

    Strips outlooks ("/ stable", "/ negative"), watch markers ("*-", "*+",
    "(neg)"), parentheticals, and surrounding whitespace. Returns the base
    token uppercased for S&P-style ratings and CamelCase for Moody's-style
    ratings (matching RATING_SCALE keys), or None if the input is empty or
    explicitly says no-rating.

        parse_rating('BBB+ *- / watch-') -> 'BBB+'
        parse_rating('Baa3 / stable')    -> 'Baa3'
        parse_rating('N.A. / n.a.')      -> None
        parse_rating('')                 -> None
    """
    if raw is None:
        return None
    s = str(raw).strip()
    if not s:
        return None
    # Strip outlook tail: split on '/' and take left part
    s = s.split('/', 1)[0].strip()
    # Strip parentheticals
    if '(' in s:
        s = s.split('(', 1)[0].strip()
    # Strip watch markers (*-, *+) and standalone '*'
    for tok in ('*-', '*+', '*'):
        s = s.replace(tok, '')
    s = s.strip()
    if not s:
        return None
    # Normalize whitespace
    s = ' '.join(s.split())
    if s.upper() in NO_RATING_TOKENS:
        return None
    # Match against scale: try as-is, then S&P uppercase, then Moody's CamelCase
    if s in RATING_SCALE:
        return s
    if s.upper() in RATING_SCALE:
        return s.upper()
    # Try Moody's-style: first letter uppercase, rest lowercase preserving digits
    cap = s[0].upper() + s[1:].lower() if len(s) > 1 else s.upper()
    if cap in RATING_SCALE:
        return cap
    return None  # unknown / unparseable


def rating_tier(token):
    """Numeric tier for a normalized rating token; None if unknown."""
    if token is None:
        return None
    return RATING_SCALE.get(token)


def is_subordinated_bond(bond):
    """True if a bond record represents subordinated debt (incl. perpetual,
    CoCo, hybrid). Case-insensitive on both Covered Type and FOType."""
    ct = (bond.get('Covered Type') or '').strip().upper()
    fo = (bond.get('FOType') or '').strip().lower()
    if ct in SUBORDINATED_COVER_TYPES:
        return True
    return any(kw in fo for kw in SUBORDINATED_FOTYPE_KEYWORDS)


def format_int(raw) -> str:
    try:
        v = int(float(raw))
        return f'{v:,}'
    except (TypeError, ValueError):
        return (raw or '').strip()


def format_date(raw) -> str:
    """Accept common date shapes and return dd.mm.yyyy."""
    s = (raw or '').strip()
    for fmt in ('%d.%m.%Y', '%d/%m/%Y', '%Y-%m-%d', '%Y/%m/%d',
                '%m/%d/%Y', '%d-%m-%Y'):
        try:
            return datetime.strptime(s, fmt).strftime('%d.%m.%Y')
        except ValueError:
            pass
    return s

def _parse_maturity_date(raw):
    """Return a datetime for sorting, or None if unparseable."""
    s = (raw or '').strip()
    if not s:
        return None
    for fmt in ('%d.%m.%Y', '%d/%m/%Y', '%Y-%m-%d', '%Y/%m/%d',
                '%m/%d/%Y', '%d-%m-%Y'):
        try:
            return datetime.strptime(s, fmt)
        except ValueError:
            pass
    return None


# ══════════════════════════════════════════════════════════════════════════════
# 6. DATA MODEL
# ══════════════════════════════════════════════════════════════════════════════

class GEMData:
    """Central data access with pre-computed indexes and enriched display fields."""

    def __init__(self, paths: dict):
        print('[data] loading bonds…');            self.bonds          = load_bonds(paths['bond_data'])
        print('[data] loading issuers…');          self.issuers        = load_issuers(paths['issuer_data'])
        print('[data] loading bond updates…');     self.bond_updates   = load_bond_updates(paths['bond_update'])
        print('[data] loading issuer updates…');   self.issuer_updates = load_issuer_updates(paths['issuer_update'])
        print('[data] loading color flags…');      self.color_flags    = load_color_flags(paths['color_flags'])
        print('[data] loading issuer texts…');     self.issuer_texts   = load_issuer_texts(paths['issuer_texts'])
        print('[data] loading ratings…');          self.issuer_ratings = load_issuer_ratings(paths['issuer_ratings'])

        # Optional prior week for Changes/Additions/Deletions
        self.prev_bond_updates = {}
        self.prev_bonds = {}
        if paths.get('prev_bond_data') and paths.get('prev_bond_update'):
            prev_all = {}
            for r in _rows_csv(paths['prev_bond_data'], quotechar='"'):
                isin = (r.get('Isin') or '').strip()
                if isin:
                    prev_all[isin] = r
            for r in _rows_csv(paths['prev_bond_update'], delimiter=';'):
                isin = (r.get('ISIN') or '').strip()
                if isin and (r.get('TopListCategory') or '').strip() == 'GEM':
                    self.prev_bond_updates[isin] = r
                    if isin in prev_all:
                        self.prev_bonds[isin] = prev_all[isin]
            # Apply the SAME eligibility rule to previous-week data so the
            # week-on-week diff is apples-to-apples. This is the structural
            # guarantee that prevents tightened rules from showing up as
            # spurious 'Deletions' relative to a previously broader policy:
            # both weeks pass through self._classify_for_list, so any rule
            # change there auto-applies symmetrically.
            kept_prev_isins = set()
            for isin, b in self.prev_bonds.items():
               upd = self.prev_bond_updates.get(isin, {})
               decision = self._classify_for_list(b, upd)
               if decision['eligible']:
                   kept_prev_isins.add(isin)
               elif decision['reason'] == 'near_maturity':
                   # Bonds dropped because they're approaching maturity
                   # SHOULD show up as Deletions this week — keep them in
                   # prev_bonds so the week-on-week diff catches them.
                   kept_prev_isins.add(isin)
            self.prev_bonds = {i: b for i, b in self.prev_bonds.items()
                               if i in kept_prev_isins}
            self.prev_bond_updates = {i: u for i, u in self.prev_bond_updates.items()
                                       if i in kept_prev_isins}

        self._build_indexes()
        # Audit reports populated during _apply_subordinated_rule and
        # _build_ratings_consistency_report (called from _build_indexes).
        # The CLI `main()` writes them out as CSV next to the PDF.
        print(f'[data] {len(self.em_bonds):,} EM bonds identified '
              f'(after subordinated-HY exclusion)')
        if self.subordinated_rule_report:
            n_excl = sum(1 for r in self.subordinated_rule_report
                         if r['decision'] == 'excluded')
            n_flag = sum(1 for r in self.subordinated_rule_report
                         if r['decision'] == 'flagged_for_review')
            print(f'[audit] subordinated bonds: '
                  f'{len(self.subordinated_rule_report)} reviewed, '
                  f'{n_excl} excluded (HY), {n_flag} flagged for review')
        if self.ratings_consistency_report:
            n_diff = sum(1 for r in self.ratings_consistency_report
                         if r['discrepancy'])
            print(f'[audit] ratings consistency: '
                  f'{n_diff} bonds where computed IG/HY disagrees with WMRFlag')

    def _build_indexes(self):
        self.bond_by_isin = {b.get('Isin', '').strip(): b for b in self.bonds
                             if (b.get('Isin') or '').strip()}
        self.bonds_by_gk = defaultdict(list)
        for b in self.bonds:
            gk = (b.get('GK_Nummer') or '').strip()
            if gk:
                self.bonds_by_gk[gk].append(b)

        # Filter the universe through the SAME classifier the prev-week diff
        # uses. Subordinated bonds are also recorded in the audit report.
        self.em_bonds = []
        self.subordinated_rule_report = []
        for b in self.bonds:
            isin = (b.get('Isin') or '').strip()
            upd  = self.bond_updates.get(isin, {})
            decision = self._classify_for_list(b, upd)
            if decision['reason'] == 'not_gem':
                continue
            if decision['reason'] in ('subordinated_ig', 'subordinated_hy',
                                       'subordinated_no_flag',
                                       'subordinated_unrated'):
                self.subordinated_rule_report.append(
                    self._build_subordinated_audit_row(b, decision))
            if decision['eligible']:
                self.em_bonds.append(b)

        # Build the ratings-consistency report against the kept set
        self.ratings_consistency_report = (
            self._build_ratings_consistency_report(self.em_bonds))

    # --------------------------------------------------------- LIST ELIGIBILITY
    #
    # Single source of truth for "does this bond belong on the EM Bond List?".
    # Used for both current-week filtering AND previous-week diff filtering.
    # If you tighten or loosen the inclusion rules, change them HERE — both
    # week's datasets will adjust together, so a newly-enforced rule will
    # never advertise itself by appearing in the Deletions section.

    def _classify_for_list(self, bond, bond_update=None):
        """Decide whether one bond record is eligible for the EM Bond List.

        Returns a dict:
            eligible : bool
            reason   : one of
                'not_emma'              — bond is outside the EMMA region
                'eligible'              — non-subordinated EM bond, kept
                'subordinated_ig'       — subordinated, issuer IG, kept
                'subordinated_hy'       — subordinated, issuer HY, kept
                'subordinated_no_flag'  — subordinated, WMRFlag missing or
                                           unknown — excluded (never silently
                                           include)

        The decision authority is the analyst-adjusted WMRFlag because it
        reflects a third rating agency that isn't in the CSV inputs.
        """
        if bond_update is None:
            isin = (bond.get('Isin') or '').strip()
            bond_update = self.bond_updates.get(isin, {})
        if (bond_update or {}).get('TopListCategory', '').strip() != 'GEM':
            return {'eligible': False, 'reason': 'not_gem'}
        maturity = (bond.get('Maturity') or '').strip()
        if maturity:
           mat_date = None
           for fmt in ('%d.%m.%Y', '%Y-%m-%d', '%d/%m/%Y'):
              try:
                  mat_date = datetime.strptime(maturity, fmt).date()
                  break
              except ValueError:
                  continue
           if mat_date and (mat_date - date.today()).days < 180:
              # Venezuelan sovereigns are defaulted — keep them on the list
              # until a restructuring happens, regardless of stated maturity.
              gk = (bond.get('GK_Nummer') or '').strip()
              is_ve_sov = (self.issuer_type(gk) == 'SOV'
                           and self.issuer_country_code(gk) == 'VE')
              if not is_ve_sov:
                  return {'eligible': False, 'reason': 'near_maturity'}
              
        if not is_subordinated_bond(bond):
            return {'eligible': True, 'reason': 'eligible'}
        # WMRFlag lives in BOTH the bond-data and bond-update files. Some
        # weeks one is empty while the other is populated — consult both
        # before declaring the flag missing.
        wmr = ((bond.get('WMRFlag') or '').strip().upper() or
               (bond_update.get('WMRFlag') or '').strip().upper())
        if wmr.startswith('HY'):
            # NOW INCLUDE subordinated HY bonds
            return {'eligible': True, 'reason': 'subordinated_hy'}
        if wmr.startswith('IG'):
            return {'eligible': True, 'reason': 'subordinated_ig'}
        # WMRFlag missing or unrecognized → fall back to a COMPUTED IG/HY
        # decision derived from the worst-of S&P/Moody's bond rating. This
        # stops a transient data-quality gap (flag empty one week, populated
        # the next) from making a bond churn between 'no_flag/excluded' and
        # 'eligible/included' — which otherwise surfaces as a spurious
        # Addition or Deletion in the week-on-week diff.
        gk = (bond.get('GK_Nummer') or '').strip()
        computed = self.effective_issuer_rating(gk, bond, bond_update)
        if computed['is_ig'] is True:
            return {'eligible': True, 'reason': 'subordinated_ig'}
        if computed['is_ig'] is False:
            return {'eligible': True, 'reason': 'subordinated_hy'}
        # Neither WMRFlag nor any agency rating is available — common for
        # Asian REIT perps and other genuinely unrated sub instruments.
        # The published EMBL includes these bonds (their inclusion is
        # editorial, not rating-driven), so we INCLUDE rather than exclude.
        # Route to HY for grade-bucket placement (unrated sub debt is more
        # naturally grouped with speculative-grade than investment-grade).
        # The audit report still surfaces the reason so reviewers can spot
        # the implicit HY classification.
        return {'eligible': True, 'reason': 'subordinated_unrated'}

    def _build_subordinated_audit_row(self, bond, decision):
        """Convert a (bond, classifier-decision) pair into one row of the
        subordinated_rule_report. Pulled out so the classifier itself stays
        a pure decision function with no dependency on display formatting.
        """
        isin = (bond.get('Isin') or '').strip()
        gk   = (bond.get('GK_Nummer') or '').strip()
        r    = self.effective_issuer_rating(gk, bond)
        ct   = (bond.get('Covered Type') or '').strip().upper()
        fo   = (bond.get('FOType') or '').strip()
        wmr  = (bond.get('WMRFlag') or '').strip().upper()
        issuer_name = self.issuer_display_name(gk, fallback=bond.get('IssuerName', ''))
        reason = decision['reason']
        decision_label = {
            'subordinated_ig':       'included',
            'subordinated_hy':       'included',
            'subordinated_unrated':  'included_as_HY',
            'subordinated_no_flag':  'flagged_for_review',
        }[reason]
        reason_text = {
            'subordinated_ig':
                f'subordinated, issuer IG per WMRFlag={wmr}',
            'subordinated_hy':
                f'subordinated, issuer HY per WMRFlag={wmr}',
            'subordinated_unrated':
                f'subordinated bond with no WMRFlag and no parseable agency '
                f'rating — included with implicit HY grade. Review whether '
                f'the published EMBL classifies this issuer differently.',
            'subordinated_no_flag':
                f'subordinated bond with missing/unknown WMRFlag={wmr!r}; '
                f'cannot determine IG/HY',
        }[reason]
        return {
            'isin': isin, 'gk': gk,
            'issuer': issuer_name,
            'cover_type': ct, 'fo_type': fo,
            'sp_raw': r['sp_raw'] or '', 'mdy_raw': r['mdy_raw'] or '',
            'sp_token': r['sp_token'] or '', 'mdy_token': r['mdy_token'] or '',
            'worst_tier': '' if r['worst_tier'] is None else r['worst_tier'],
            'rating_status': r['status'],
            'wmr_flag': wmr,
            'decision': decision_label,
            'reason': reason_text,
        }

    # --------------------------------------------------------- ratings helpers

    def effective_issuer_rating(self, gk, bond=None, bond_update=None):
        """Resolve the effective S&P and Moody's ratings for a BOND, then
        classify IG vs HY using the WORSE (higher-tier-number) of the two.

        Lookup precedence (most-specific source first):
          • S&P:    bond_update.RatingSP  > bond.SP > issuer_ratings[gk].SP
                  > issuer.SPIssuerRating
          • Moody's: bond_update.RatingMdy > issuer_ratings[gk].MDY
                  > issuer.MDYIssuerRating > bond.MDY

        bond_update (PublishableBondDataUpdate.txt) carries BOND-level agency
        ratings. For subordinated bonds those are typically notched DOWN one
        or more rungs from the issuer rating — preferring them when present
        is what makes the displayed rating match the actual bond rating
        (rather than the issuer rating, which would over-state credit
        quality for sub debt).

        Returns a dict:
            sp_raw, mdy_raw           – raw strings as found (may be None)
            sp_token, mdy_token       – normalized tokens (may be None)
            sp_tier, mdy_tier         – ints in [0..21] or None
            worst_tier                – max(sp_tier, mdy_tier), ignoring None
            is_ig                     – True/False/None (None = unclassifiable)
            status                    – 'both' | 'sp_only' | 'mdy_only' |
                                         'none' | 'unparseable'
        """
        ir = self.issuer_ratings.get(gk, {}) if gk else {}
        issuer = self.issuer_record(gk) if gk else {}
        bond = bond or {}
        # When the caller didn't provide one, look up the bond's update row
        # by ISIN — keeps existing call sites that pass only (gk, bond) working.
        if bond_update is None:
            isin = (bond.get('Isin') or '').strip()
            bond_update = self.bond_updates.get(isin, {}) if isin else {}
        bond_update = bond_update or {}

        sp_raw  = ((bond_update.get('RatingSP') or '').strip() or
                   (bond.get('SP') or '').strip() or
                   (ir.get('SP') or '').strip() or
                   (issuer.get('SPIssuerRating') or '').strip() or None)
        mdy_raw = ((bond_update.get('RatingMdy') or '').strip() or
                   (ir.get('MDY') or '').strip() or
                   (issuer.get('MDYIssuerRating') or '').strip() or
                   (bond.get('MDY') or '').strip() or None)

        sp_token  = parse_rating(sp_raw)
        mdy_token = parse_rating(mdy_raw)
        sp_tier   = rating_tier(sp_token)
        mdy_tier  = rating_tier(mdy_token)

        present = [t for t in (sp_tier, mdy_tier) if t is not None]
        if not present:
            if sp_raw or mdy_raw:
                status = 'unparseable'   # values exist but neither could be normalized
            else:
                status = 'none'
            worst, is_ig = None, None
        else:
            worst = max(present)
            is_ig = (worst <= IG_MAX_TIER)
            if sp_tier is not None and mdy_tier is not None:
                status = 'both'
            elif sp_tier is not None:
                status = 'sp_only'
            else:
                status = 'mdy_only'

        return {
            'sp_raw': sp_raw, 'mdy_raw': mdy_raw,
            'sp_token': sp_token, 'mdy_token': mdy_token,
            'sp_tier': sp_tier, 'mdy_tier': mdy_tier,
            'worst_tier': worst, 'is_ig': is_ig,
            'status': status,
        }

    # --------------------------------------------------------- audit & filters

    def _build_ratings_consistency_report(self, bonds):
        """Compare the WMRFlag in source against an IG/HY classification
        derived from the lowest available agency rating. Returns a list of
        per-bond audit rows, with `discrepancy=True` when the two disagree.
        """
        rows = []
        for b in bonds:
            isin = (b.get('Isin') or '').strip()
            gk   = (b.get('GK_Nummer') or '').strip()
            r    = self.effective_issuer_rating(gk, b)
            wmr  = (b.get('WMRFlag') or '').strip().upper()
            # WMRFlag baseline: anything starting with HY counts as HY
            # (covers HY, HY*-, HY*+ etc.); IG-* otherwise; '' = unknown
            if not wmr:
                source_label = 'unknown'
            elif wmr.startswith('HY'):
                source_label = 'HY'
            else:
                source_label = 'IG'
            if r['is_ig'] is True:
                computed_label = 'IG'
            elif r['is_ig'] is False:
                computed_label = 'HY'
            else:
                computed_label = 'unclassifiable'

            discrepancy = (
                computed_label in ('IG', 'HY') and source_label in ('IG', 'HY')
                and computed_label != source_label
            )
            rows.append({
                'isin': isin, 'gk': gk,
                'issuer': self.issuer_display_name(gk, fallback=b.get('IssuerName', '')),
                'sp_raw': r['sp_raw'] or '', 'mdy_raw': r['mdy_raw'] or '',
                'sp_token': r['sp_token'] or '', 'mdy_token': r['mdy_token'] or '',
                'sp_tier': '' if r['sp_tier'] is None else r['sp_tier'],
                'mdy_tier': '' if r['mdy_tier'] is None else r['mdy_tier'],
                'worst_tier': '' if r['worst_tier'] is None else r['worst_tier'],
                'computed_grade': computed_label,
                'source_wmr_flag': wmr,
                'source_grade': source_label,
                'discrepancy': discrepancy,
            })
        return rows

    # ------------------------------------------------------------------ resolvers

    def issuer_record(self, gk):
        """Merge issuer master + weekly update for a GK."""
        base = dict(self.issuers.get(gk, {}))
        base.update({k: v for k, v in (self.issuer_updates.get(gk, {}) or {}).items() if v})
        return base

    def issuer_display_name(self, gk, fallback=''):
        """CTL_List_Name → CML_Publish_Name → fallback.
        Apply Title Case to any source that comes through as ALL CAPS."""
        rec = self.issuer_record(gk)
        for field in ('CTL_List_Name', 'CML_Publish_Name'):
            value = (rec.get(field) or '').strip()
            if value:
                # If the source is ALL CAPS (likely the Access export format), title-case it.
                if value.isupper() or (value == value.upper() and any(c.isalpha() for c in value)):
                    return title_case_name(value)
                return value
        return (fallback or '').strip()

    def issuer_country_code(self, gk):
        rec = self.issuer_record(gk)
        return (rec.get('WMR_Country') or '').strip().upper()

    def issuer_type(self, gk):
        rec = self.issuer_record(gk)
        return (rec.get('WMR_IssuerType') or '').strip().upper()

    def analyst_for_gk(self, gk):
        rec = self.issuer_record(gk)
        a = (rec.get('IssuerAnalyst') or rec.get('Analyst') or '').strip()
        return analyst_name(a)

    def issuer_wmr_rating(self, gk):
        rec = self.issuer_record(gk)
        return (rec.get('WMR_Rating') or '').strip() or 'NR'

    def issuer_trend(self, gk):
        rec = self.issuer_record(gk)
        return (rec.get('WMR_Trend') or '').strip() or 'Stable'

    def issuer_description_text(self, gk):
        rec = self.issuer_record(gk)
        desc = (rec.get('Issuer_Description') or '').strip()
        if desc:
            return desc
        # Fallback to IssuerTexts.txt
        for t in self.issuer_texts.get(gk, []):
            d = (t.get('IssuerDescription') or '').strip()
            if d:
                return d
        return ''

    # ------------------------------------------------------------------ bond row

    def bond_row(self, b):
        """Return a normalized dict suitable for rendering a single bond row."""
        isin   = (b.get('Isin') or '').strip()
        valor  = (b.get('Valor') or '').strip()
        gk     = (b.get('GK_Nummer') or '').strip()
        upd    = self.bond_updates.get(isin, {})

        # Issuer display
        itype  = self.issuer_type(gk)
        cc     = self.issuer_country_code(gk)
        name   = self.issuer_display_name(gk, fallback=b.get('IssuerName', ''))
        if itype in ('SOV', 'SUPRA'):
            display = name                           # sovereigns: no parens
        else:
            display = f'{name} ({country_name(cc)})' if cc else name

        ccy   = (b.get('CCY') or '').strip().upper()
        coupon= format_percent(b.get('Coupon'))
        matur = format_date(b.get('Maturity')) or 'Perpetual'
        matur_dt = _parse_maturity_date(b.get('Maturity')) # for chronological sort
        px    = format_price(b.get('PXASK_ExecDesk'))
        # Yield-to-maturity is meaningless for floating-rate bonds — the
        # ExecDesk feed often returns 0.0 for them. Detect FRNs (variable
        # coupon or 'float' in FOType) and emit 'n/a' instead so readers
        # don't take a literal 0% YTM as a market signal.
        cpn_type_raw = (b.get('CpnType') or '').strip().lower()
        fo_type_raw  = (b.get('FOType')  or '').strip().lower()
        is_floater   = (cpn_type_raw in ('variable', 'fixed/variable') or
                        'float' in fo_type_raw)
        yld   = 'n/a' if is_floater else format_price(b.get('YLDASK_ExecDesk'))
        # Resolve S&P / Moody's via bond_update (BOND-level) + bond + issuer
        # precedence. Passing `upd` makes sub bonds show the notched-down
        # bond rating from PublishableBondDataUpdate rather than the issuer
        # rating, which is what the published list expects.
        eff   = self.effective_issuer_rating(gk, b, upd)
        sp    = eff['sp_token'] or 'n/a'
        mdy   = eff['mdy_token'] or 'n/a'
        rating= f'{sp} / {mdy}'   # "S&P / Moody's"
        min_a = format_int(b.get('MinAmt'))
        min_i = format_int(b.get('MinInc'))
        min_d = f'{min_a} / {min_i}' if min_i else min_a

        green = GREEN_LABELS.get((b.get('GreenBond') or '').strip().upper(), '')

        # ---- Restrictions column ------------------------------------------
        # Derived via rules in the CONFIG block. Plain-vanilla senior bullet
        # bonds (redeemable=N, retractable=N, SEN, FOType='straight bond')
        # have no embedded derivatives → no restrictions applied.
        rest_parts = []

        cover_type   = (b.get('Covered Type') or '').strip().upper()
        redeemable   = (b.get('redeemable')   or '').strip().upper()
        retractable  = (b.get('retractable')  or '').strip().upper()
        fo_type      = (b.get('FOType')       or '').strip().lower()

        has_complex_cover   = cover_type in MIFID_COMPLEX_COVER_TYPES
        has_embedded_option = (redeemable == 'Y' or retractable == 'Y')
        has_complex_fotype  = any(kw in fo_type for kw in MIFID_COMPLEX_FOTYPE_KEYWORDS)

        is_complex = has_complex_cover or has_embedded_option or has_complex_fotype

        # Rule 1: MiFID "Complex bond"
        if is_complex:
            rest_parts.append('1')

        # Rule 2: PRIIPs KID missing. For EM bonds, non-EEA issuers don't
        # register KIDs, so restriction 2 co-applies with restriction 1 by
        # default. Override per-ISIN via KID_MISSING_ISINS if needed.
        if PRIIPS_KID_AUTO_APPLY:
            if is_complex:
                rest_parts.append('2')
        elif isin in KID_MISSING_ISINS:
            rest_parts.append('2')

        restrictions = ', '.join(rest_parts)

        # IG/HY grade: use the classifier's decision so unrated sub bonds
        # (subordinated_unrated, subordinated_hy) end up in the speculative
        # bucket rather than silently defaulting to IG when WMRFlag is
        # empty. Senior bonds still trust WMRFlag with bond_update fallback.
        decision = self._classify_for_list(b, upd)
        sub_reason_to_grade = {
            'subordinated_ig':      'Investment grade issuers',
            'subordinated_hy':      'Speculative grade issuers',
            'subordinated_unrated': 'Speculative grade issuers',
        }
        if decision['reason'] in sub_reason_to_grade:
            grade = sub_reason_to_grade[decision['reason']]
        else:
            wmr_flag = ((b.get('WMRFlag') or '').strip().upper() or
                        (upd.get('WMRFlag') or '').strip().upper())
            if wmr_flag.startswith('HY'):
                grade = 'Speculative grade issuers'
            elif wmr_flag.startswith('IG'):
                grade = 'Investment grade issuers'
            else:
                # No WMRFlag for a senior bond — derive from agency ratings.
                grade = ('Investment grade issuers' if eff['is_ig'] is True
                         else 'Speculative grade issuers' if eff['is_ig'] is False
                         else 'Investment grade issuers')
        wmr_flag = (b.get('WMRFlag') or '').strip().upper()

        rec_code = (upd.get('WMR_Bond_Recommendation') or b.get('WMR_Bond_Recommendation') or '').strip().upper()
        view = {'OP': 'attr.', 'UP': 'exp.', 'SELL': 'Sell'}.get(rec_code, 'fair')

        return {
            'isin': isin,
            'valor': valor,
            'gk': gk,
            'isin_valor': f'{isin} / {valor}' if valor else isin,
            'issuer': display,
            'issuer_raw': name,
            'country': cc,
            'country_display': country_name(cc),
            'itype': itype,
            'ccy': ccy,
            'coupon': coupon,
            'maturity': matur,
            'maturity_date': matur_dt, #for chronological sort
            'offer_price': px,
            'offer_yield': yld,
            'ratings': rating,
            'min_denom': min_d,
            'green': green,
            'restrictions': restrictions,
            'comment': (b.get('WMR_Bond_Comment') or '').strip(),
            'view': view,
            'rec': rec_code,
            'region': effective_region(name, cc),
            'grade': grade,
            'wmr_flag': wmr_flag,
        }

    # ------------------------------------------------------------------ selectors

    def top_list_bonds(self):
       out = []
       for b in self.em_bonds:
           # Defensive: re-check eligibility at render time.
           if not self._classify_for_list(b)['eligible']:
               continue
           pu  = (b.get('Product_Use') or '').strip()
           rec = (b.get('WMR_Bond_Recommendation') or '').strip()
           if pu == '7' and rec != 'Sell':
               out.append(b)
       return out

    def sell_list_bonds(self):
       return [b for b in self.em_bonds
               if (b.get('WMR_Bond_Recommendation') or '').strip() == 'Sell'
               and self._classify_for_list(b)['eligible']]

    # ------------------------------------------------------------------ changes
    # Severity ranking used to classify week-on-week recommendation flips.
    # Lower = better. Sell is worst (off the active list).
    _REC_SEVERITY = {'OP': 0, 'MP': 1, '': 1, 'UP': 2, 'SELL': 3}
    _REC_VIEW = {'OP': 'attr.', 'MP': 'fair', '': 'fair', 'UP': 'exp.', 'SELL': 'exp.'}

    def _rec_code(self, upd, bond):
        return (
            (upd or {}).get('WMR_Bond_Recommendation') or
            (bond or {}).get('WMR_Bond_Recommendation') or
            ''
        ).strip().upper()

    def _change_row(self, isin, bond, view_prior, view_new):
        """Build a single row dict for the Changes table from a bond record.

        Used for both current-week (Upgrades/Downgrades/Additions) and
        previous-week (Deletions) bonds. Resolves issuer display name from
        the current issuer master (an issuer that disappears entirely is rare).
        """
        gk    = (bond.get('GK_Nummer') or '').strip()
        itype = self.issuer_type(gk)
        cc    = self.issuer_country_code(gk)
        name  = self.issuer_display_name(gk, fallback=bond.get('IssuerName', ''))
        if itype in ('SOV', 'SUPRA'):
            issuer = name
        else:
            issuer = f'{name} ({country_name(cc)})' if cc else name
        return {
            'isin':       isin,
            'issuer':     issuer,
            'ccy':        (bond.get('CCY') or '').strip().upper(),
            'coupon':     format_percent(bond.get('Coupon')),
            'maturity':   format_date(bond.get('Maturity')),
            'view_prior': view_prior,
            'view_new':   view_new,
        }

    def recommendation_changes(self):
        """Diff this week vs. previous week to produce four lists for the
        'Changes to the recommendations' page:

          - upgrades:   ISIN on both weeks, severity decreased
                        (e.g. fair → attr., exp. → fair, Sell → MP/OP)
          - downgrades: ISIN on both weeks, severity increased
                        (e.g. attr. → fair, fair → exp., MP → Sell)
          - additions:  ISIN present this week, absent last week
          - deletions:  ISIN present last week, absent this week

        Each row is a dict with isin, issuer, ccy, coupon, maturity,
        view_prior, view_new (None where the bond didn't exist that week).
        Returns {} if no previous-week data is available.
        """
        if not self.prev_bond_updates:
            return {'upgrades': [], 'downgrades': [], 'additions': [], 'deletions': []}

        # Index current EMMA bonds by ISIN
        cur = {}
        for b in self.em_bonds:
            isin = (b.get('Isin') or '').strip()
            if isin:
                cur[isin] = (self._rec_code(self.bond_updates.get(isin, {}), b), b)
       
        # Previous-week index already filtered to EMMA in __init__
        prev = {}
        for isin, upd in self.prev_bond_updates.items():
            b = self.prev_bonds.get(isin, {})
            prev[isin] = (self._rec_code(upd, b), b)
            
        upgrades, downgrades, additions, deletions = [], [], [], []
        for isin in set(cur) | set(prev):
            c = cur.get(isin)
            p = prev.get(isin)
            if c and not p:
                additions.append(self._change_row(
                    isin, c[1], None, self._REC_VIEW.get(c[0], 'fair')))
            elif p and not c:
                deletions.append(self._change_row(
                    isin, p[1], self._REC_VIEW.get(p[0], 'fair'), None))
            elif c and p:
                c_sev = self._REC_SEVERITY.get(c[0], 1)
                p_sev = self._REC_SEVERITY.get(p[0], 1)
                if c_sev == p_sev:
                    continue
                row = self._change_row(
                    isin, c[1],
                    self._REC_VIEW.get(p[0], 'fair'),
                    self._REC_VIEW.get(c[0], 'fair'),
                )
                (upgrades if c_sev < p_sev else downgrades).append(row)
        
        sort_key = lambda r: ((r['issuer'] or '').lower(), r['isin'])
        return {
            'upgrades':   sorted(upgrades,   key=sort_key),
            'downgrades': sorted(downgrades, key=sort_key),
            'additions':  sorted(additions,  key=sort_key),
            'deletions':  sorted(deletions,  key=sort_key),
        }

    def has_recommendation_changes(self):
        c = self.recommendation_changes()
        return any(c.values())

    def reference_list_bonds(self):
       out = []
       for b in self.em_bonds:
           # Defensive: re-check eligibility at render time. Guarantees
           # mutual exclusion with the Deletions list — any bond classified
           # as ineligible (e.g. near_maturity) cannot land here.
           if not self._classify_for_list(b)['eligible']:
               continue
           rec = (b.get('WMR_Bond_Recommendation') or '').strip()
           ccy = (b.get('CCY') or '').strip().upper()
           if rec == 'Sell':              continue   # Sell bonds live on the Sell Recommendations page
           # NOTE: Top-List bonds (Product_Use=='7') ARE included here. The
           # Top List is an overlay/highlight of the analyst's top picks —
           # the published list still expects those same bonds to appear in
           # their currency/region reference table alongside their peers.
           if not ccy:                    continue   # skip bonds with no currency
           if ccy in EXCLUDED_CURRENCIES: continue   # skip excluded currencies
           out.append(b)
       return out

    def data_timestamp(self):
        """Return the publication date as 'D MMM YYYY' (e.g. '5 May 2026')."""
        dt = None
        if self.bonds:
            tk = (self.bonds[0].get('TimeKey') or '').strip()
            for fmt in ('%d.%m.%Y %H:%M:%S', '%Y-%m-%d %H:%M:%S'):
                try:
                    dt = datetime.strptime(tk, fmt)
                    break
                except ValueError:
                    continue
        if dt is None:
            dt = datetime.now()
        return f"{dt.day} {dt.strftime('%b %Y')}"


    # ------------------------------------------------------------------ sorting

    @staticmethod
    def sort_key(row):
        mat = row.get('maturity_date')
        mat_sort = mat if mat is not None else datetime.max
        return (
            row['region'],
            row['grade'],
            ISSUER_TYPE_RANK.get(row['itype'], 3),
            (row['issuer_raw'] or '').lower(),
            mat_sort,
        )


# ══════════════════════════════════════════════════════════════════════════════
# 7. FLOWABLES — bookmark / section-title / source
# ══════════════════════════════════════════════════════════════════════════════

class BookmarkAnchor(Flowable):
    """Zero-height flowable that plants a bookmark at the current Y on the page."""
    def __init__(self, key):
        Flowable.__init__(self)
        self.key = key
        self.width = 0
        self.height = 0
    def draw(self):
        self.canv.bookmarkPage(self.key)
        # also register as an outline entry? optional

class SetSectionTitle(ActionFlowable):
    """Instructs the doc template to change its running header title/subtitle."""
    def __init__(self, title, subtitle=''):
        ActionFlowable.__init__(self)
        self.title = title
        self.subtitle = subtitle
    def apply(self, doc):
        doc.current_title = self.title
        doc.current_subtitle = self.subtitle

class SetSource(ActionFlowable):
    """Instructs the doc template to change the centered source footer text."""
    def __init__(self, text):
        ActionFlowable.__init__(self)
        self.text = text
    def apply(self, doc):
        doc.current_source = self.text

class RotatedText(Flowable):
    """Draws a text string rotated 90° counter-clockwise. Sized to fit inside
    a given width/height box — the caller is responsible for allocating the
    table cell or container it lives in.
    """
    def __init__(self, text, font_name='Helvetica-Bold', font_size=11,
                 color=colors.black, width=10, height=100):
        Flowable.__init__(self)
        self.text = text
        self.font_name = font_name
        self.font_size = font_size
        self.color = color
        self.width = width
        self.height = height

    def wrap(self, aw, ah):
        return self.width, self.height

    def draw(self):
        c = self.canv
        c.saveState()
        c.translate(self.width / 2, self.height / 2)
        c.rotate(90)
        c.setFont(self.font_name, self.font_size)
        c.setFillColor(self.color)
        c.drawCentredString(0, -self.font_size * 0.35, self.text)
        c.restoreState()


class TOCEntryMarker(Flowable):
    """Records a TOC entry's page number as the flow lays out."""
    def __init__(self, key):
        Flowable.__init__(self)
        self.key = key
        self.width = 0
        self.height = 0
    def draw(self):
        # Page number captured in doc template's afterFlowable
        pass


# ══════════════════════════════════════════════════════════════════════════════
# 8. DOC TEMPLATE
# ══════════════════════════════════════════════════════════════════════════════

class GEMDocTemplate(BaseDocTemplate):
    """BaseDocTemplate with running title/subtitle, source, and page X of Y."""

    def __init__(self, filename, logo_path=None, **kw):
        BaseDocTemplate.__init__(self, filename, **kw)
        self.logo_path        = logo_path
        self.total_pages      = 0
        self.current_title    = REPORT_TITLE
        self.current_subtitle = REPORT_SUBTITLE
        self.current_source   = SOURCE_DEFAULT
        self.page_tracker     = {}     # bookmark_key → page_number
        self.publication_date = ''
        self.disclosures_page_placeholder = 'x'

        frame = Frame(
            MARGIN_LEFT, MARGIN_BOTTOM,
            CONTENT_WIDTH, CONTENT_HEIGHT,
            leftPadding=0, rightPadding=0, topPadding=0, bottomPadding=0,
            id='content',
        )
        self.addPageTemplates([
            PageTemplate(id='First', frames=[frame], onPage=self._draw_first_page),
            PageTemplate(id='Later', frames=[frame], onPage=self._draw_later_page),
            PageTemplate(id='LaterClean', frames=[frame],
                         onPage=self._draw_later_clean_page),
        ])

    # ------------------------------------------------------------- callbacks

    def afterFlowable(self, flowable):
        if isinstance(flowable, BookmarkAnchor):
            self.page_tracker[flowable.key] = self.page
        elif isinstance(flowable, TOCEntryMarker):
            self.page_tracker[flowable.key] = self.page

    def _draw_first_page(self, canv, doc):
        canv.saveState()
        # Logo top-left
        if self.logo_path and os.path.exists(self.logo_path):
            try:
                canv.drawImage(self.logo_path, MARGIN_LEFT, PAGE_HEIGHT - 1.9*cm,
                               width=2.2*cm, height=1.1*cm,
                               preserveAspectRatio=True, mask='auto')
            except Exception as e:
                print(f'[warn] could not place logo: {e}')
        # Publication date + CIO GWM block top-right
        canv.setFont('Helvetica', 8)
        canv.setFillColor(UBS_MID)
        y = PAGE_HEIGHT - 1.0*cm
        for line in (
            f'Publication date: {self.publication_date}',
            'Chief Investment Office GWM',
            'Investment Research',
            'For investors outside of the US',
        ):
            canv.drawRightString(PAGE_WIDTH - MARGIN_RIGHT, y, line)
            y -= 0.35*cm
        # Footer: source + page number
        # On page 1 the Source line is rendered in the cover body; skip it
        # in the footer so it doesn't appear twice.
        self._draw_footer(canv, doc, include_disclaimer=True, skip_source=True)
        canv.restoreState()

    def _draw_later_page(self, canv, doc):
        self._draw_later_page_impl(canv, doc, skip_indicative=False)

    def _draw_later_clean_page(self, canv, doc):
        """Same chrome as Later, but without the '* Indicative values'
        footnote on the left. Used for the Rating Definitions and
        Sanctions Notice reference pages, which have no bond-pricing
        context where the indicative-values disclaimer would apply.
        """
        self._draw_later_page_impl(canv, doc, skip_indicative=True)

    def _draw_later_page_impl(self, canv, doc, skip_indicative=False):
        canv.saveState()
        # Title + subtitle band
        canv.setFillColor(UBS_DARK)
        canv.setFont('Helvetica-Bold', 13)
        canv.drawString(MARGIN_LEFT, PAGE_HEIGHT - 1.0*cm, self.current_title)
        if self.current_subtitle:
            canv.setFont('Helvetica', 8)
            canv.setFillColor(UBS_MID)
            canv.drawString(MARGIN_LEFT, PAGE_HEIGHT - 1.5*cm, self.current_subtitle)
        # Thin rule
        canv.setStrokeColor(UBS_HEADER_BG)
        canv.setLineWidth(0.6)
        canv.line(MARGIN_LEFT, PAGE_HEIGHT - 1.8*cm,
                  PAGE_WIDTH - MARGIN_RIGHT, PAGE_HEIGHT - 1.8*cm)
        self._draw_footer(canv, doc, skip_indicative=skip_indicative)
        canv.restoreState()

    def _draw_footer(self, canv, doc, include_disclaimer=False, skip_source=False,
                      skip_indicative=False):
        canv.setFont('Helvetica', 7.5)
        canv.setFillColor(UBS_MID)
       
        # '* Indicative values' footnote — sits on the left on every later
        # page so it's always visible alongside the Offer price* / Offer yield*
        # column headers that live in the table header above. Suppressed on
        # reference/notice pages that have nothing to do with bond pricing.
        if not include_disclaimer and not skip_indicative:
            canv.drawString(MARGIN_LEFT, FOOTER_Y, INDICATIVE_FOOTNOTE)
        # Source (centered) — suppressed on cover page where it's in the body
        if not skip_source:
            canv.drawCentredString(PAGE_WIDTH / 2.0, FOOTER_Y, self.current_source)
        # Page X of Y (right)
        if self.total_pages:
            canv.drawRightString(PAGE_WIDTH - MARGIN_RIGHT, FOOTER_Y,
                                 f'Page {doc.page} of {self.total_pages}')
        else:
            canv.drawRightString(PAGE_WIDTH - MARGIN_RIGHT, FOOTER_Y,
                                 f'Page {doc.page}')


# ══════════════════════════════════════════════════════════════════════════════
# 9. STYLES
# ══════════════════════════════════════════════════════════════════════════════

def build_styles(fonts):
    """Create all ParagraphStyles using the resolved font family."""
    ss = getSampleStyleSheet()
    f = fonts
    def P(name, **kw):
        base = ParagraphStyle(name, parent=ss['Normal'])
        for k, v in kw.items():
            setattr(base, k, v)
        return base
    return {
        'cover_title':  P('cover_title', fontName=f['bold'],   fontSize=24, leading=28, textColor=colors.black, spaceAfter=2),
        'cover_sub':    P('cover_sub',   fontName=f['light'],  fontSize=11, leading=14, textColor=UBS_MID,     spaceAfter=14),
        'h1':           P('h1',          fontName=f['bold'],   fontSize=16, leading=20, textColor=UBS_DARK,    spaceAfter=6),
        'h2':           P('h2',          fontName=f['bold'],   fontSize=12, leading=15, textColor=UBS_DARK,    spaceBefore=8, spaceAfter=3),
        'h3':           P('h3',          fontName=f['bold'],   fontSize=10, leading=12, textColor=UBS_DARK,    spaceBefore=4, spaceAfter=2),
        'body':         P('body',        fontName=f['light'],  fontSize=9,  leading=12, textColor=UBS_DARK,    spaceAfter=3),
        'body_sm':      P('body_sm',     fontName=f['light'],  fontSize=8,  leading=10, textColor=UBS_DARK,    spaceAfter=2),
        'body_it':      P('body_it',     fontName=f['italic'], fontSize=9,  leading=12, textColor=UBS_DARK),
        'table_hdr':    P('table_hdr',   fontName=f['bold'],   fontSize=7.5,leading=9,  textColor=colors.white, alignment=TA_LEFT),
        'table_hdr_c':  P('table_hdr_c', fontName=f['bold'],   fontSize=7.5,leading=9,  textColor=colors.white, alignment=TA_CENTER),
        'cell':         P('cell',        fontName=f['light'],  fontSize=7,  leading=8.5),
        'cell_b':       P('cell_b',      fontName=f['bold'],   fontSize=7,  leading=8.5),
        'cell_right':   P('cell_right',  fontName=f['light'],  fontSize=7,  leading=8.5, alignment=TA_RIGHT),
        'comment':      P('comment',     fontName=f['italic'], fontSize=6.5,leading=8,  textColor=UBS_MID, leftIndent=8*mm),
        'footnote':     P('footnote',    fontName=f['light'],  fontSize=7,  leading=9,  textColor=UBS_MID, spaceBefore=2),
        'group_hdr':    P('group_hdr',   fontName=f['bold'],   fontSize=8,  leading=10, textColor=UBS_DARK),
        'group_hdr_r':  P('group_hdr_r', fontName=f['italic'], fontSize=8,  leading=10, textColor=UBS_MID, alignment=TA_RIGHT),
        'toc_a':        P('toc_a',       fontName=f['bold'],   fontSize=7,  leading=8.5, textColor=UBS_DARK),
        'toc_b':        P('toc_b',       fontName=f['light'],  fontSize=7,  leading=8.5, textColor=UBS_DARK, leftIndent=3*mm),
        'toc_page':     P('toc_page',    fontName=f['light'],  fontSize=7,  leading=8.5, textColor=UBS_DARK, alignment=TA_RIGHT),
        'spec_warn':    P('spec_warn',   fontName=f['italic'], fontSize=8,  leading=10, textColor=UBS_DARK, backColor=UBS_LIGHT,
                          leftIndent=4, rightIndent=4, spaceBefore=2, spaceAfter=3),
    }


# ══════════════════════════════════════════════════════════════════════════════
# 10. PDF BUILDER
# ══════════════════════════════════════════════════════════════════════════════

class GEMPDFBuilder:
    """Orchestrates the two-pass build. Pass 1 collects bookmarks / page count;
    pass 2 writes the final PDF with accurate TOC and 'Page X of Y' numbers."""

    # Column layout for bond-list tables (widths sum to CONTENT_WIDTH)
    BOND_COL_WIDTHS = [
        14*mm,  # View (attr. / fair / exp.)  — analyst's relative-value call
        36*mm,  # ISIN / Valor
        52*mm,  # Issuer
        14*mm,  # Currency
        14*mm,  # Coupon
        18*mm,  # Maturity
        15*mm,  # Offer price¹
        15*mm,  # Offer yield¹
        24*mm,  # Ratings
        28*mm,  # Min denom / increment
        22*mm,  # Green / Social / Sust
        25*mm,  # Restrictions
    ]
    BOND_HEADERS = [
        'View',
        'ISIN / Valor', 'Issuer', 'Currency', 'Coupon', 'Maturity',
        f'Offer\nprice{INDICATIVE_MARK}',
        f'Offer\nyield{INDICATIVE_MARK}',
        'Ratings\n(S&P / Moody\'s)',
        'Minimum denomination /\nincrement',
        'Green, Social,\nSustainability',
        'Restrictions',
    ]

    def __init__(self, data: GEMData, output: str, logo_path=None):
        self.data   = data
        self.output = output
        self.logo_path = logo_path
        self.fonts  = register_fonts()
        self.styles = build_styles(self.fonts)
        self.story  = []
        self.page_tracker = {}
        self.total_pages  = 0
        # TOC sections declared during build
        self.toc_sections = []  # list of (bookmark_key, display_name, indent)
        # Capture the moment this PDF is being generated — used for the
        # 'Market data shown … is as of …' line on the cover page.
        now = datetime.now()
        self.pdf_created_at = now
        self.pdf_created_label = f'{now.strftime("%m %d %Y, %H:%M")} {PUBLICATION_TZ_LABEL}'

    # ------------------------------------------------------------- build driver

    def build(self):
        # Pass 1: build story (cover-page TOC will be empty since bookmarks unknown)
        # and capture bookmark page numbers + the complete toc_sections list.
        print('[pdf] pass 1 (collect page numbers)…')
        self._build_story()
        pass1_toc = list(self.toc_sections)  # preserve for pass 2 cover page

        import tempfile
        tmp = tempfile.NamedTemporaryFile(suffix='.pdf', delete=False); tmp.close()
        doc = GEMDocTemplate(
            tmp.name, logo_path=self.logo_path,
            pagesize=landscape(A4),
            leftMargin=MARGIN_LEFT, rightMargin=MARGIN_RIGHT,
            topMargin=MARGIN_TOP, bottomMargin=MARGIN_BOTTOM,
        )
        doc.publication_date = self.data.data_timestamp()
        doc.build(self.story)
        self.page_tracker = dict(doc.page_tracker)
        self.total_pages  = doc.page
        print(f'[pdf]   collected {len(self.page_tracker)} bookmarks, {self.total_pages} pages')
        os.unlink(tmp.name)

        # Pass 2: prepopulate toc_sections so the cover page can render it with
        # correct page numbers from page_tracker.
        print('[pdf] pass 2 (final with accurate TOC)…')
        self._prepopulated_toc = pass1_toc
        self._build_story()
        doc = GEMDocTemplate(
            self.output, logo_path=self.logo_path,
            pagesize=landscape(A4),
            leftMargin=MARGIN_LEFT, rightMargin=MARGIN_RIGHT,
            topMargin=MARGIN_TOP, bottomMargin=MARGIN_BOTTOM,
        )
        doc.publication_date = self.data.data_timestamp()
        doc.total_pages = self.total_pages
        doc.page_tracker = dict(self.page_tracker)
        doc.build(self.story)
        print(f'[pdf] → {self.output}')

    # ------------------------------------------------------------- story build

    def _build_story(self):
        self.story = []
        # If pass 2, seed toc_sections so the cover page renders a filled TOC.
        # Otherwise start empty and let _register populate it as sections build.
        if getattr(self, '_prepopulated_toc', None):
            self.toc_sections = list(self._prepopulated_toc)
            self._toc_seeded = True
        else:
            self.toc_sections = []
            self._toc_seeded = False
        self._add_cover_page()
        self.story.append(NextPageTemplate('Later'))

        # IMPORTANT: at every section boundary, emit SetSectionTitle BEFORE the
        # PageBreak. ActionFlowable.apply() runs when the flowable is processed,
        # so if SetSectionTitle comes before the break, the NEW page's onPage
        # callback (which reads doc.current_title) will see the new title.
        # Emitting it AFTER the break leaks the previous section's header onto
        # the first page of the new section.

        self._begin_section('guidance_bonds',
                            'Guidance on bond characteristics',
                            'Guidance on bond characteristics', indent=0)
        self._add_guidance_bonds()

        self._begin_section('guidance_ratings',
                            'Guidance on credit ratings and subordinated bonds',
                            'Guidance on credit ratings and subordinated bonds', indent=0)
        self._add_guidance_ratings()

        self._begin_section('toplist',
                            'Top Emerging Markets Bond List',
                            'Top Emerging Markets Bond List', indent=0)
        self._add_top_list_section()

        if self.data.sell_list_bonds():
            self._begin_section('sell',
                                'Sell Recommendations',
                                'Sell Recommendations', indent=0)
            self._add_sell_list_section()

        # ---- Changes to the recommendations (week-on-week diff) --------------
        if self.data.has_recommendation_changes():
           self._begin_section('changes',
                               'Changes to the recommendations',
                               'Changes to the recommendations', indent=0)
           self._add_changes_section()
           self._begin_section('additions_deletions',
                               'Additions and Deletions',
                               'Additions and Deletions', indent=0)
           self._add_additions_deletions_section()

        # ---- Reference Lists (Bonds by Currency × Region × Grade) ------------
        # No dedicated 'Reference lists' cover page — the first sub-section
        # (e.g. 'Bonds in USD, Asia') is the band's first page. We still want
        # a top-level 'Reference lists' TOC row pointing there; it's registered
        # on the first iteration inside _add_reference_lists so the bookmark
        # and the hyperlink resolve to the same page.
        self._add_reference_lists()

        # ---- Issuer Descriptions & Appendices --------------------------------
        self._begin_section('issuerdesc',
                            'Issuer descriptions and credit-risk flags',
                            'Issuer descriptions and credit-risk flags', indent=0)
        self._add_issuer_descriptions()

        self._begin_section('disclosures_start', 'Appendix', 'Appendix', indent=0)
        self._add_appendix()

        # Final reference pages — switch to the 'LaterClean' template so the
        # left-side '* Indicative values' footnote is suppressed (it doesn't
        # apply to either of these pages — no bond pricing involved).
        self.story.append(NextPageTemplate('LaterClean'))

        self._begin_section('rating_definitions',
                            'Rating Definitions',
                            'Rating Definitions', indent=0)
        self._add_rating_definitions()

        self._begin_section('sanctions_notice',
                            'Sanctions notice',
                            'Sanctions notice', indent=0)
        self._add_sanctions_notice()

    def _register(self, key, toc_label, indent):
        """Plant a bookmark at the current position and record a TOC entry.
        Does NOT change the running-header title — callers must emit
        SetSectionTitle before the preceding PageBreak."""
        self.story.append(BookmarkAnchor(key))
        if not getattr(self, '_toc_seeded', False):
            self.toc_sections.append((key, toc_label, indent))

    def _begin_section(self, key, toc_label, section_title, subtitle='',
                       indent=0, page_break=True):
        """Start a new section correctly:

          1. Queue SetSectionTitle so it is processed BEFORE the PageBreak.
             This means the NEW page's onPage() callback will read the new
             title when drawing the running header.
          2. Emit PageBreak (unless caller opts out).
          3. Plant the bookmark (on the new page).
          4. Add the TOC entry.

        No h1 Paragraph is emitted here — the section builder is responsible
        for rendering its own body title if it wants to.
        """
        if subtitle == '':
            subtitle = REPORT_SUBTITLE
        # 1) Set title BEFORE the break
        self.story.append(SetSectionTitle(section_title, subtitle))
        # 2) Page break
        if page_break:
            self.story.append(PageBreak())
        # 3) Bookmark on the new page
        self.story.append(BookmarkAnchor(key))
        # 4) TOC entry
        if not getattr(self, '_toc_seeded', False):
            self.toc_sections.append((key, toc_label, indent))

    # ------------------------------------------------------------- COVER PAGE

    def _add_cover_page(self):
        s = self.styles
        # Title
        self.story.append(Paragraph(REPORT_TITLE, s['cover_title']))
        self.story.append(Paragraph(REPORT_SUBTITLE, s['cover_sub']))
        self.story.append(Spacer(1, 4*mm))

        # Two-column layout: TOC on left, Analysts on right
        left_col = self._build_toc_column()
        right_col = self._build_analysts_column()

        two_col = Table(
            [[left_col, right_col]],
            colWidths=[CONTENT_WIDTH * 0.55, CONTENT_WIDTH * 0.45],
        )
        two_col.setStyle(TableStyle([
            ('VALIGN', (0, 0), (-1, -1), 'TOP'),
            ('LEFTPADDING',  (0, 0), (-1, -1), 0),
            ('RIGHTPADDING', (0, 0), (-1, -1), 0),
            ('TOPPADDING',   (0, 0), (-1, -1), 0),
            ('BOTTOMPADDING',(0, 0), (-1, -1), 0),
        ]))
        self.story.append(two_col)

        # ---- Market-data note with the PDF creation timestamp ---------------
        # Rendered as part of the cover body (NOT in the page footer) so the
        # reader sees it alongside the analyst list.
        self.story.append(Spacer(1, 6*mm))
        note_sty = ParagraphStyle('CoverNote', parent=s['body_sm'],
                                   fontSize=8, leading=11, textColor=UBS_DARK)
        self.story.append(Paragraph(
            MARKET_DATA_NOTE.format(timestamp=self.pdf_created_label),
            note_sty,
        ))

        # Source line — bolded, same position as the published cover
        self.story.append(Spacer(1, 3*mm))
        src_sty = ParagraphStyle('CoverSource', parent=s['body_sm'],
                                  fontSize=8, leading=10,
                                  textColor=UBS_DARK, fontName=self.fonts['bold'])
        self.story.append(Paragraph(
            '<b>Source:</b> IPS CFMP Fixed Income Execution Desk (internal: goto/bondpricing)',
            src_sty,
        ))

        # The analyst-certification disclaimer + page number stays in the
        # page-1 footer (see _draw_first_page).

    def _build_toc_column(self):
        s = self.styles
        # Header
        header_row = [Paragraph('<u><b>Table of contents</b></u>', s['h3']), '']
        entry_rows = []
        for key, label, indent in self.toc_sections:
            p_num = self.page_tracker.get(key, '…')
            style = s['toc_b'] if indent else s['toc_a']
            link = f'<a href="#{key}" color="#0563C1"><u>{label}</u></a>'
            entry_rows.append([Paragraph(link, style), Paragraph(str(p_num), s['toc_page'])])

        col1_w = CONTENT_WIDTH * 0.55 - 12*mm
        col_widths = [col1_w - 12*mm, 12*mm]
        tbl = Table([header_row] + entry_rows, colWidths=col_widths)
        tbl.setStyle(TableStyle([
            ('VALIGN', (0, 0), (-1, -1), 'TOP'),
            ('LEFTPADDING',  (0, 0), (-1, -1), 0),
            ('RIGHTPADDING', (0, 0), (-1, -1), 0),
            ('TOPPADDING',   (0, 0), (-1, -1), 0),
            ('BOTTOMPADDING',(0, 0), (-1, -1), 0),
        ]))
        return tbl

    def _build_analysts_column(self):
        """Render the analysts' expertise column from the curated ANALYST_ROSTER.

        The roster is hand-maintained in the CONFIG block at the top of the file.
        Each entry renders as two stacked paragraphs on the left (name in bold,
        title/branch beneath in a smaller gray) and the expertise on the right.
        """
        s = self.styles

        # Header bar — gray band across the full column width
        hdr = Table(
            [[Paragraph('<b>Analysts\u2019 area of expertise</b>', s['h3'])]],
            colWidths=[CONTENT_WIDTH * 0.45 - 2*mm],
        )
        hdr.setStyle(TableStyle([
            ('BACKGROUND',    (0, 0), (-1, -1), UBS_LIGHT),
            ('LEFTPADDING',   (0, 0), (-1, -1), 4),
            ('TOPPADDING',    (0, 0), (-1, -1), 3),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 3),
        ]))

        # Styles local to the analyst rows
        title_sty = ParagraphStyle('ana_title', parent=s['body_sm'],
                                    fontSize=6.5, leading=8,
                                    textColor=colors.HexColor('#666666'))
        name_sty  = ParagraphStyle('ana_name', parent=s['cell_b'],
                                    fontSize=7.5, leading=9, textColor=UBS_DARK)
        expertise_sty = ParagraphStyle('ana_exp', parent=s['body_sm'],
                                        fontSize=7.5, leading=9,
                                        textColor=UBS_DARK, alignment=TA_RIGHT)

        rows = []
        for name, title, expertise in ANALYST_ROSTER:
            left = [Paragraph(f'<b>{name}</b>', name_sty),
                    Paragraph(title, title_sty)]
            right = Paragraph(expertise, expertise_sty)
            rows.append([left, right])

        col_w = CONTENT_WIDTH * 0.45 - 2*mm
        tbl = Table(rows, colWidths=[col_w * 0.52, col_w * 0.48])
        tbl.setStyle(TableStyle([
            ('VALIGN',         (0, 0), (-1, -1), 'TOP'),
            ('LEFTPADDING',    (0, 0), (-1, -1), 4),
            ('RIGHTPADDING',   (0, 0), (-1, -1), 4),
            ('TOPPADDING',     (0, 0), (-1, -1), 2),
            ('BOTTOMPADDING',  (0, 0), (-1, -1), 2),
        ]))

        # Stack header + rows
        stack = Table(
            [[hdr], [tbl]],
            colWidths=[CONTENT_WIDTH * 0.45 - 2*mm],
        )
        stack.setStyle(TableStyle([
            ('VALIGN',        (0, 0), (-1, -1), 'TOP'),
            ('LEFTPADDING',   (0, 0), (-1, -1), 0),
            ('RIGHTPADDING',  (0, 0), (-1, -1), 0),
            ('TOPPADDING',    (0, 0), (-1, -1), 0),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 0),
        ]))
        return stack

    # ------------------------------------------------------------- GUIDANCE PAGES

    def _add_guidance_bonds(self):
        s = self.styles
        self.story.append(SetSource(SOURCE_RATINGS))

        # Guidance-page column header: omit 'Issuer' (it has no special meaning
        # that needs explanation); include 'View' at the front.
        g_headers = [
            'View', 'ISIN / Valor', 'Currency', 'Coupon', 'Maturity',
            f'Offer\nprice{INDICATIVE_MARK}',
            f'Offer\nyield{INDICATIVE_MARK}',
            'Ratings\n(S&P / Moody\'s)',
            'Minimum denomination /\nincrement',
            'Green, Social,\nSustainability',
            'Restrictions',
        ]
        # Widths sum to CONTENT_WIDTH (≈277 mm landscape A4 with 1 cm margins)
        g_widths = [15*mm, 38*mm, 20*mm, 18*mm, 22*mm, 20*mm, 20*mm,
                    28*mm, 38*mm, 30*mm, 28*mm]
        total = sum(g_widths)
        if total > CONTENT_WIDTH:
            scale = CONTENT_WIDTH / total
            g_widths = [x * scale for x in g_widths]
        hdr = [Paragraph(h.replace('&', '&amp;'), s['table_hdr']) for h in g_headers]
        tbl = Table([hdr], colWidths=g_widths, repeatRows=1)
        tbl.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, -1), UBS_HEADER_BG),
            ('TEXTCOLOR',  (0, 0), (-1, -1), colors.white),
            ('VALIGN',     (0, 0), (-1, -1), 'MIDDLE'),
            ('LEFTPADDING',(0, 0), (-1, -1), 3),
            ('RIGHTPADDING',(0, 0), (-1, -1), 3),
            ('TOPPADDING', (0, 0), (-1, -1), 5),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 5),
        ]))
        self.story.append(tbl)
        self.story.append(Spacer(1, 4*mm))

        sections = [
            ('View',
             'Our relative-value assessment of the bond, based on bottom-up credit analysis combined with '
             'top-down tactical calls:<br/>'
             '&nbsp;&nbsp;<b>attr.</b> (Attractive) — Bonds seen as attractive are expected to generate a total return exceeding the average return of comparable instruments. Our recommendation can stem from a positive view on the issuer’s credit profile not fully reflected in the price, unduly high risk premiums, call probability, the risk of coupon deferrals, and external factors including regulatory intervention.<br/>'
             '&nbsp;&nbsp;<b>fair</b> (Fair) — Bonds seen as fair are expected to produce a total return broadly in line with the average return of comparable instruments.<br/>'
             '&nbsp;&nbsp;<b>exp.</b> (Expensive) — Bonds seen as expensive are expected to earn a total return that is less than the average return of comparable instruments. Our recommendation can stem from a negative view on the issuer’s credit profile not fully reflected in the price, unduly tight risk premiums, call probability, the risk of coupon deferrals, and external factors including regulatory intervention.<br/>'
             '&nbsp;&nbsp;<b>exp.</b> (Sell) — A sell recommendation is assigned when the risk of an adverse outcome for an instrument exceeds what is reflected in its current valuation. Such situations can include those in which the instrument appears likely to post negative total returns until redemption, either due to a highly negative yield to maturity or an imminent call at a price below market valuations.<br/>'),
            ('Offer yield',
             'The offer yield refers to the yield-to-maturity measure. Please note that the displayed values are '
             'indicative values only.'),
            ('Green, Social, Sustainability',
             'An entry in this column signifies that the bond is a green, social, or sustainability bond. '
             'The comment section will specify which category the bond falls under.<br/>'
             '<b>Green bond:</b> Pursues projects that contribute to environmental sustainability, resulting in '
             'better access to long-term project financing.<br/>'
             '<b>Social bond:</b> Intended for funding projects that address or mitigate a specific social issue '
             'or seek to achieve positive social outcomes.<br/>'
             '<b>Sustainability bond:</b> Used to finance projects that may have a primarily green or social '
             'purpose, but also exhibit significant benefits from the other category. Issuers may also finance '
             'a blend of green and social projects with sustainability bonds.'),
            ('Restrictions',
             'Includes restrictions for European Economic Area domiciled investors:<br/>'
             '&nbsp;&nbsp;1) Complex bond under MiFID<br/>'
             '&nbsp;&nbsp;2) PRIIPS relevant bond, KID missing<br/>'
             'Please refer to our education note "Understanding Bonds" published on 18 December 2018.'),
        ]
        for title, body in sections:
            bar = Table([[Paragraph(f'<b>{title}</b>', s['table_hdr'])]],
                        colWidths=[CONTENT_WIDTH])
            bar.setStyle(TableStyle([
                ('BACKGROUND', (0,0), (-1,-1), UBS_HEADER_BG),
                ('TEXTCOLOR',  (0,0), (-1,-1), colors.white),
                ('LEFTPADDING',(0,0), (-1,-1), 4),
                ('TOPPADDING', (0,0), (-1,-1), 3),
                ('BOTTOMPADDING', (0,0), (-1,-1), 3),
            ]))
            self.story.append(bar)
            self.story.append(Spacer(1, 1*mm))
            self.story.append(Paragraph(body, s['body']))
            self.story.append(Spacer(1, 3*mm))

    def _add_guidance_ratings(self):
        s = self.styles
        self.story.append(SetSource(SOURCE_DEFAULT))

        bar = Table([[Paragraph('<b>Credit rating definitions</b>', s['table_hdr'])]],
                    colWidths=[CONTENT_WIDTH])
        bar.setStyle(TableStyle([
            ('BACKGROUND', (0,0), (-1,-1), UBS_HEADER_BG),
            ('TEXTCOLOR',  (0,0), (-1,-1), colors.white),
            ('LEFTPADDING',(0,0), (-1,-1), 4),
            ('TOPPADDING', (0,0), (-1,-1), 3),
            ('BOTTOMPADDING', (0,0), (-1,-1), 3),
        ]))
        self.story.append(bar)
        self.story.append(Spacer(1, 2*mm))

        # Two-column IG vs SG definitions
        ig = [
            ('AAA / Aaa', 'Issuer / Bonds have exceptionally strong credit quality. AAA is the best credit quality.'),
            ('AA+ / Aa1, AA / Aa2, AA- / Aa3', 'Issuer / Bonds have very strong credit quality.'),
            ('A+ / A1, A / A2, A- / A3', 'Issuer / Bonds have high credit quality.'),
            ('BBB+ / Baa1, BBB / Baa2, BBB- / Baa3',
             'Issuer / Bonds have adequate credit quality. This is the lowest Investment Grade category.'),
        ]
        sg = [
            ('BB+ / Ba1, BB / Ba2, BB- / Ba3',
             'Issuer / Bonds have weak credit quality. This is the highest Speculative Grade category.'),
            ('B+ / B1, B / B2, B- / B3', 'Issuer / Bonds have very weak credit quality.'),
            ('CCC+ / Caa1, CCC / Caa2, CCC- / Caa3',
             'Issuer / Bonds have extremely weak credit quality.'),
            ('CC / Ca, C / -', 'Issuer / Bonds have very high risk of default.'),
            ('D / C', 'Obligor failed to make payment on one or more of its financial commitments.'),
        ]
        def rating_col(title, rows):
            data = [[Paragraph(f'<b>{title}</b>', s['h3'])]]
            for grade, desc in rows:
                data.append([Paragraph(f'<b>{grade}</b><br/>{desc}', s['body_sm'])])
            t = Table(data, colWidths=[CONTENT_WIDTH / 2 - 2*mm])
            t.setStyle(TableStyle([
                ('VALIGN', (0,0), (-1,-1), 'TOP'),
                ('LEFTPADDING',(0,0), (-1,-1), 0),
                ('RIGHTPADDING',(0,0), (-1,-1), 0),
                ('TOPPADDING', (0,0), (-1,-1), 2),
                ('BOTTOMPADDING', (0,0), (-1,-1), 2),
            ]))
            return t
        two = Table(
            [[rating_col('Investment Grade', ig), rating_col('Speculative Grade', sg)]],
            colWidths=[CONTENT_WIDTH / 2, CONTENT_WIDTH / 2],
        )
        two.setStyle(TableStyle([
            ('VALIGN', (0,0), (-1,-1), 'TOP'),
            ('LEFTPADDING',(0,0), (-1,-1), 0),
            ('RIGHTPADDING',(0,0), (-1,-1), 4*mm),
        ]))
        self.story.append(two)
        self.story.append(Spacer(1, 3*mm))
        self.story.append(Paragraph(
            'Issuer ratings may differ between rating agencies. Analysts may choose to assign the lowest rating '
            'instead of an average rating. This may lead to a situation in which issuers with an average investment '
            'grade rating appear in the sub-investment grade section of the Emerging Markets Bond List.',
            s['body_sm']))

        # --- Subordinated bonds block ---------------------------------------
        self.story.append(Spacer(1, 5*mm))

        sub_bar = Table([[Paragraph('<b>Subordinated bonds</b>', s['table_hdr'])]],
                        colWidths=[CONTENT_WIDTH])
        sub_bar.setStyle(TableStyle([
            ('BACKGROUND',    (0, 0), (-1, -1), UBS_HEADER_BG),
            ('TEXTCOLOR',     (0, 0), (-1, -1), colors.white),
            ('LEFTPADDING',   (0, 0), (-1, -1), 4),
            ('TOPPADDING',    (0, 0), (-1, -1), 3),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 3),
        ]))
        self.story.append(sub_bar)
        self.story.append(Spacer(1, 2*mm))

        # Left: narrative prose
        sub_narrative = [
            Paragraph(
                'If a bond issuer were to default, a subordinated bond would rank lower in status than other '
                'debt when it comes to a claim on the company\u2019s assets. This makes subordinated bonds '
                'riskier than higher ranked bonds.',
                s['body_sm']),
            Paragraph(
                'In addition, such bonds might become less liquid during periods of adverse market conditions '
                'than higher ranked instruments, making it more difficult to sell such bonds during period of '
                'higher financial market volatility.',
                s['body_sm']),
            Paragraph(
                'Moreover, we include subordinated bonds issued by issuers rated '
                '\u2018Speculative Grade\u2019 (see definition above) on the list.',
                s['body_sm']),
        ]
        left_cell = Table([[p] for p in sub_narrative],
                          colWidths=[CONTENT_WIDTH / 2 - 2*mm])
        left_cell.setStyle(TableStyle([
            ('VALIGN',       (0, 0), (-1, -1), 'TOP'),
            ('LEFTPADDING',  (0, 0), (-1, -1), 0),
            ('RIGHTPADDING', (0, 0), (-1, -1), 0),
            ('TOPPADDING',   (0, 0), (-1, -1), 0),
            ('BOTTOMPADDING',(0, 0), (-1, -1), 2),
        ]))

        # Right: light-gray sub-header + 3-row label/description tier table
        right_col_w = CONTENT_WIDTH / 2 - 2*mm
        tier_label_w = 25*mm
        tier_desc_w  = right_col_w - tier_label_w

        tier_subhdr = Table(
            [[Paragraph(
                '<b>Subordinated debt is divided into 2 main tiers. Tier 1 debt is subordinate to Tier 2 debt.</b>',
                s['body_sm'])]],
            colWidths=[right_col_w])
        tier_subhdr.setStyle(TableStyle([
            ('BACKGROUND',    (0, 0), (-1, -1), UBS_LIGHT),
            ('LEFTPADDING',   (0, 0), (-1, -1), 4),
            ('RIGHTPADDING',  (0, 0), (-1, -1), 4),
            ('TOPPADDING',    (0, 0), (-1, -1), 3),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 3),
        ]))

        tier_rows = [
            ('Tier 1',
             'The maturity of Tier 1 debt is perpetual, however, the issuer has the right to call the bond at '
             'the earliest after five years, then at each coupon date. Calling the bond is only possible if '
             'sufficient funds are available for repayment. Interest can be paid on a fixed or floating basis, '
             'the bond is not collateralised nor guaranteed.'),
            ('Upper Tier 2',
             'Upper Tier 2 debt is perpetual, and its coupons are deferrable and cumulative, interest and '
             'principal can be written down.'),
            ('Lower Tier 2',
             'Lower Tier 2 debt has a fixed maturity of at least 5 years and interest payments may only be '
             'suspended in the case of bankruptcy.'),
        ]
        tier_data = []
        for label, desc in tier_rows:
            tier_data.append([
                Paragraph(f'<b>{label}</b>', s['body_sm']),
                Paragraph(desc, s['body_sm']),
            ])
        tier_tbl = Table(tier_data, colWidths=[tier_label_w, tier_desc_w])
        tier_tbl.setStyle(TableStyle([
            ('VALIGN',        (0, 0), (-1, -1), 'TOP'),
            ('LEFTPADDING',   (0, 0), (-1, -1), 4),
            ('RIGHTPADDING',  (0, 0), (-1, -1), 4),
            ('TOPPADDING',    (0, 0), (-1, -1), 4),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 4),
            ('LINEBELOW',     (0, 0), (-1, -2), 0.25, UBS_HEADER_BG),
        ]))

        right_cell = Table(
            [[tier_subhdr], [tier_tbl]],
            colWidths=[right_col_w])
        right_cell.setStyle(TableStyle([
            ('VALIGN',       (0, 0), (-1, -1), 'TOP'),
            ('LEFTPADDING',  (0, 0), (-1, -1), 0),
            ('RIGHTPADDING', (0, 0), (-1, -1), 0),
            ('TOPPADDING',   (0, 0), (-1, -1), 0),
            ('BOTTOMPADDING',(0, 0), (-1, -1), 0),
        ]))

        sub_two = Table(
            [[left_cell, right_cell]],
            colWidths=[CONTENT_WIDTH / 2, CONTENT_WIDTH / 2])
        sub_two.setStyle(TableStyle([
            ('VALIGN',       (0, 0), (-1, -1), 'TOP'),
            ('LEFTPADDING',  (0, 0), (-1, -1), 0),
            ('RIGHTPADDING', (0, 0), (-1, -1), 4*mm),
        ]))
        self.story.append(sub_two)

    # ------------------------------------------------------------- BOND TABLES

    def _bond_table_header(self):
        s = self.styles
        hdr = [Paragraph(h.replace('&', '&amp;'), s['table_hdr']) for h in self.BOND_HEADERS]
        return hdr

    def _bond_row_cells(self, row):
        s = self.styles
        # View column shows the analyst's relative-value call ('attr.' / 'fair'
        # / 'exp.'). Rendered in the default cell style (no per-value colour).
        view_cell = Paragraph(row['view'], s['cell'])
        return [
            view_cell,
            Paragraph(row['isin_valor'], s['cell']),
            Paragraph(row['issuer'], s['cell']),
            Paragraph(row['ccy'], s['cell']),
            Paragraph(row['coupon'], s['cell']),
            Paragraph(row['maturity'], s['cell']),
            Paragraph(row['offer_price'], s['cell']),
            Paragraph(row['offer_yield'], s['cell']),
            Paragraph(row['ratings'], s['cell']),
            Paragraph(row['min_denom'], s['cell']),
            Paragraph(row['green'], s['cell']),
            Paragraph(row['restrictions'], s['cell']),
        ]

    def _render_bond_table(self, rows, show_region_grade_header=False,
                            show_issuer_headers=False, show_group_analyst=True):
        """Render a list of bond rows as a single bond-list table with group sub-headers.

        Parameters
          show_region_grade_header : prepend a 'Bonds in <region>, <grade>'
                                     band whenever the region/grade changes
                                     (used by Top List / Sell List).
          show_issuer_headers      : add a per-issuer header row each time the
                                     GK changes — [Issuer | Credit Outlook | Analyst].
                                     Used by the Reference Lists. Top/Sell lists
                                     leave it off (too short, would get noisy).
          show_group_analyst       : when NOT showing per-issuer headers, put a
                                     modal 'Analyst: …' on the sub-group header.
        """
        s = self.styles

        # Sort rows
        rows = sorted(rows, key=GEMData.sort_key)

        # Group by (region, grade, issuer_type) for sub-headers; then by issuer for issuer rows
        data = [self._bond_table_header()]
        style_cmds = [
            ('BACKGROUND', (0, 0), (-1, 0), UBS_HEADER_BG),
            ('TEXTCOLOR',  (0, 0), (-1, 0), colors.white),
            ('VALIGN',     (0, 0), (-1, -1), 'MIDDLE'),
            ('LEFTPADDING',(0, 0), (-1, -1), 3),
            ('RIGHTPADDING',(0, 0), (-1, -1), 3),
            ('TOPPADDING', (0, 0), (-1, -1), 2),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 2),
            ('ROWBACKGROUNDS', (0, 1), (-1, -1), [colors.white, UBS_LIGHT]),
            ('LINEBELOW', (0, 0), (-1, 0), 0.25, colors.white),
        ]

        current_region_grade = None
        current_itype = None
        current_gk = None

        # Styles for the per-issuer header row
        issuer_hdr_style = ParagraphStyle(
            'issuer_hdr', parent=s['group_hdr'],
            fontSize=8.5, leading=10, textColor=UBS_DARK)
        outlook_hdr_style = ParagraphStyle(
            'outlook_hdr', parent=s['group_hdr'],
            fontSize=8, leading=10, textColor=UBS_DARK,
            fontName=self.fonts['light'], alignment=TA_CENTER)
        analyst_hdr_style = ParagraphStyle(
            'issuer_analyst_hdr', parent=s['group_hdr_r'],
            fontSize=8, leading=10, textColor=UBS_MID,
            fontName=self.fonts['italic'], alignment=TA_RIGHT)

        for row in rows:
            rg = (row['region'], row['grade'])
            it = row['itype']

            if show_region_grade_header and rg != current_region_grade:
                # Section header: e.g. 'Bonds in Asia, Investment grade issuers'
                label = f'Bonds in {row["region"]}, {row["grade"]}'
                data.append([Paragraph(f'<b>{label}</b>', s['group_hdr'])] + [''] * (len(self.BOND_HEADERS) - 1))
                style_cmds.append(('BACKGROUND', (0, len(data) - 1), (-1, len(data) - 1), UBS_SECTION_BG))
                style_cmds.append(('SPAN', (0, len(data) - 1), (-1, len(data) - 1)))

                # Speculative legend
                if row['grade'].startswith('Speculative'):
                    data.append([Paragraph(f'<i>{SPECULATIVE_LEGEND}</i>', s['spec_warn'])] + [''] * (len(self.BOND_HEADERS) - 1))
                    style_cmds.append(('BACKGROUND', (0, len(data) - 1), (-1, len(data) - 1), UBS_LIGHT))
                    style_cmds.append(('SPAN', (0, len(data) - 1), (-1, len(data) - 1)))
                current_region_grade = rg
                current_itype = None
                current_gk = None
            # Sub-group header: SOV/SUPRA are bucketed together as "Sovereign
            # issuers", everything else as "Corporate issuers and financials".
            # Emit a new sub-group header row whenever that bucket changes.
            it_group = 'SOV' if it in ('SOV', 'SUPRA') else 'CORP_FIN'
            if it_group != current_itype:
                group_label = 'Sovereign issuers' if it in ('SOV', 'SUPRA') else 'Corporate issuers and financials'
                n_cols = len(self.BOND_HEADERS)
                if show_issuer_headers:
                    # Per-issuer headers will carry the accurate analyst for
                    # each issuer below — don't show a (potentially conflicting)
                    # modal analyst at the itype level.
                    label_cell = Paragraph(f'<b>{group_label}</b>', s['group_hdr'])
                    row_cells = [label_cell] + [''] * (n_cols - 1)
                    data.append(row_cells)
                    style_cmds.append(('BACKGROUND', (0, len(data) - 1), (-1, len(data) - 1), UBS_SUBGROUP_BG))
                    style_cmds.append(('SPAN', (0, len(data) - 1), (-1, len(data) - 1)))
                else:
                    if show_group_analyst:
                        group_rows = [r for r in rows if (r['region'], r['grade']) == rg and r['itype'] == it]
                        analysts = [self.data.analyst_for_gk(r['gk']) for r in group_rows]
                        analysts = [a for a in analysts if a]
                        analyst = Counter(analysts).most_common(1)[0][0] if analysts else ''
                        analyst_cell = Paragraph(f'<i>Analyst: {analyst}</i>' if analyst else '', s['group_hdr_r'])
                        label_cell = Paragraph(f'<b>{group_label}</b>', s['group_hdr'])
                        row_cells = [label_cell] + [''] * (n_cols - 2) + [analyst_cell]
                        data.append(row_cells)
                        style_cmds.append(('BACKGROUND', (0, len(data) - 1), (-1, len(data) - 1), UBS_SUBGROUP_BG))
                        style_cmds.append(('SPAN', (0, len(data) - 1), (-2, len(data) - 1)))
                    else:
                        label_cell = Paragraph(f'<b>{group_label}</b>', s['group_hdr'])
                        row_cells = [label_cell] + [''] * (n_cols - 1)
                        data.append(row_cells)
                        style_cmds.append(('BACKGROUND', (0, len(data) - 1), (-1, len(data) - 1), UBS_SUBGROUP_BG))
                        style_cmds.append(('SPAN', (0, len(data) - 1), (-1, len(data) - 1)))
                current_itype = it_group
                current_gk = None

            # Per-issuer header row (Reference Lists only). Triggered when
            # the GK changes — emits one row spanning all 12 columns split
            # into three regions: issuer name | outlook | analyst.
            if show_issuer_headers and row['gk'] and row['gk'] != current_gk:
                gk = row['gk']
                outlook = self.data.issuer_trend(gk)
                analyst = self.data.analyst_for_gk(gk)
                issuer_name = row['issuer_raw'] or row['issuer']
                hdr_cells = (
                    [Paragraph(f'<b>{issuer_name}</b>', issuer_hdr_style)]
                    + [''] * 2
                    + [Paragraph(f'Credit Outlook: {outlook}', outlook_hdr_style)]
                    + [''] * 4
                    + [Paragraph(
                        f'Analyst: {analyst}' if analyst else '',
                        analyst_hdr_style)]
                    + [''] * 3
                )
                data.append(hdr_cells)
                ridx = len(data) - 1
                # Three-cell layout via SPANs:
                #   cols 0..2  → issuer name
                #   cols 3..7  → outlook (centered)
                #   cols 8..-1 → analyst (right-aligned)
                style_cmds.extend([
                    ('SPAN',       (0, ridx), (2, ridx)),
                    ('SPAN',       (3, ridx), (7, ridx)),
                    ('SPAN',       (8, ridx), (-1, ridx)),
                    ('BACKGROUND', (0, ridx), (-1, ridx), colors.white),
                    ('LINEABOVE',  (0, ridx), (-1, ridx), 0.5, UBS_HEADER_BG),
                    ('LINEBELOW',  (0, ridx), (-1, ridx), 0.25, colors.HexColor('#CCCCCC')),
                    ('TOPPADDING',    (0, ridx), (-1, ridx), 4),
                    ('BOTTOMPADDING', (0, ridx), (-1, ridx), 4),
                ])
                current_gk = gk

            data.append(self._bond_row_cells(row))

            if row['comment']:
                # Comment row spanning all columns
                data.append([Paragraph(f'<b>Comment:</b> <i>{row["comment"]}</i>', s['comment'])] +
                            [''] * (len(self.BOND_HEADERS) - 1))
                style_cmds.append(('SPAN', (0, len(data) - 1), (-1, len(data) - 1)))

        tbl = Table(data, colWidths=self.BOND_COL_WIDTHS, repeatRows=1)
        tbl.setStyle(TableStyle(style_cmds))
        self.story.append(tbl)
        # The '* Indicative values' footnote is rendered by the page-footer
        # callback (_draw_later_page), so it appears on EVERY page the table
        # spans — not only under its final row.

    # ------------------------------------------------------------- TOP LIST

    def _add_top_list_section(self):
        s = self.styles
        self.story.append(SetSource(SOURCE_DEFAULT))
        self.story.append(Paragraph(TOP_LIST_INTRO, s['body']))
        self.story.append(Spacer(1, 3*mm))

        raw_rows = [self.data.bond_row(b) for b in self.data.top_list_bonds()]
        self._render_bond_table(raw_rows, show_region_grade_header=True, show_group_analyst=False)

    def _add_sell_list_section(self):
        s = self.styles
        self.story.append(SetSource(SOURCE_DEFAULT))
        self.story.append(Paragraph(
            'A Sell recommendation is assigned when the risk of an adverse outcome for an instrument exceeds what is reflected in its current valuation. Such situations can include those in which the instrument appears likely to post negative total returns until redemption, either due to a highly negative yield to maturity or an imminent call at a price below market valuations.', s['body']))
        self.story.append(Spacer(1, 3*mm))

        raw_rows = [self.data.bond_row(b) for b in self.data.sell_list_bonds()]
        if raw_rows:
            self._render_bond_table(raw_rows, show_region_grade_header=True, show_group_analyst=False)

    # ------------------------------------------------------------- CHANGES PAGE

    _VIEW_COLOR = {
        'attr.': '#2E7D32',
        'fair':  '#4A4A4A',
        'exp.':  '#B22222',
    }

    def _view_cell(self, label):
        s = self.styles
        if not label:
            return Paragraph('', s['cell'])
        colour = self._VIEW_COLOR.get(label, '#4A4A4A')
        return Paragraph(label,
                         ParagraphStyle('view_c', parent=s['cell'], alignment=TA_CENTER))

    def _build_change_table(self, title, rows, arrow, full_width=False):
        """Build one of the four change sub-tables (Upgrades / Downgrades /
        Additions / Deletions). All four share the same column layout:

            View prior | arrow | new | ISIN | Issuer | Ccy | Coupon | Maturity

        full_width=False (default): half-width, suitable for the 2-up
                                    Upgrades/Downgrades layout.
        full_width=True:            full content width — gives Issuer column
                                    far more room and lets the table split
                                    naturally across pages (used for
                                    Additions/Deletions which can be long).
        """
        s = self.styles

        if full_width:
            table_w = CONTENT_WIDTH
        else:
            # Landscape A4: each half-table gets ~135mm. Shave 5mm off the
            # half width for breathing room vs the outer grid's right padding.
            table_w = CONTENT_WIDTH / 2 - 5*mm
        w_view_prior = 11*mm
        w_arrow      = 6*mm
        w_view_new   = 10*mm
        w_isin       = 22*mm
        w_ccy        = 11*mm
        w_coupon     = 11*mm
        w_maturity   = 15*mm
        w_issuer     = table_w - (w_view_prior + w_arrow + w_view_new +
                                   w_isin + w_ccy + w_coupon + w_maturity)
        col_widths = [w_view_prior, w_arrow, w_view_new,
                      w_isin, w_issuer, w_ccy, w_coupon, w_maturity]

        # Header bar spanning the full table width
        header_bar = [
            Paragraph(f'<b>{title}</b>', s['table_hdr_c']),
            '', '', '', '', '', '', ''
        ]
        # Sub-header row with column labels
        col_hdr_style = ParagraphStyle(
            'chg_col_hdr', parent=s['cell'], fontName=self.fonts['bold'],
            fontSize=6.5, leading=8, textColor=UBS_DARK, alignment=TA_LEFT)
        col_hdr_c = ParagraphStyle(
            'chg_col_hdr_c', parent=col_hdr_style, alignment=TA_CENTER)
        col_headers = [
            Paragraph('View<br/>prior', col_hdr_c),
            Paragraph('', col_hdr_c),
            Paragraph('new', col_hdr_c),
            Paragraph('ISIN', col_hdr_style),
            Paragraph('Issuer', col_hdr_style),
            Paragraph('Ccy', col_hdr_c),
            Paragraph('Coupon', col_hdr_c),
            Paragraph('Maturity', col_hdr_c),
        ]

        # Body rows. Wrap the arrow glyph in the symbol font (DejaVu) when
        # available so Unicode arrows (↗ ↘ + −) render correctly. If no
        # symbol font is registered, fall back to ASCII alternatives that
        # the body font (Helvetica) can render.
        arrow_style = ParagraphStyle(
           'chg_arrow', parent=s['cell'], fontSize=10, leading=11,
            textColor=UBS_DARK, alignment=TA_CENTER)
        symbol_font = self.fonts.get('symbol')
        if symbol_font:
           arrow_html = f'<font name="{symbol_font}">{arrow}</font>'
        else:
           ascii_map = {'\u2197': '+', '\u2198': '-', '\u2212': '-'}
           arrow_html = ascii_map.get(arrow, arrow)
        cell_c = ParagraphStyle('chg_cell_c', parent=s['cell'], alignment=TA_CENTER)
        body_rows = []
        for r in rows:
            body_rows.append([
                self._view_cell(r['view_prior']),
                Paragraph(arrow_html, arrow_style),
                self._view_cell(r['view_new']),
                Paragraph(r['isin'], s['cell']),
                Paragraph(r['issuer'] or '', s['cell']),
                Paragraph(r['ccy'] or '', cell_c),
                Paragraph(r['coupon'] or '', cell_c),
                Paragraph(r['maturity'] or '', cell_c),
            ])

        # Empty-state placeholder
        if not body_rows:
            empty = Paragraph(
                f'<i>No {title.lower()} this week.</i>',
                ParagraphStyle('chg_empty', parent=s['cell'],
                               textColor=UBS_MID, alignment=TA_LEFT))
            body_rows = [[empty, '', '', '', '', '', '', '']]

        data = [header_bar, col_headers] + body_rows
        tbl = Table(data, colWidths=col_widths, repeatRows=2)
        style_cmds = [
            # Title bar
            ('SPAN',          (0, 0), (-1, 0)),
            ('BACKGROUND',    (0, 0), (-1, 0), UBS_HEADER_BG),
            ('TEXTCOLOR',     (0, 0), (-1, 0), colors.white),
            ('LEFTPADDING',   (0, 0), (-1, 0), 4),
            ('TOPPADDING',    (0, 0), (-1, 0), 3),
            ('BOTTOMPADDING', (0, 0), (-1, 0), 3),
            # Column-header row
            ('BACKGROUND',    (0, 1), (-1, 1), UBS_LIGHT),
            ('TOPPADDING',    (0, 1), (-1, 1), 3),
            ('BOTTOMPADDING', (0, 1), (-1, 1), 3),
            # Body rows — tight padding so the narrow view/arrow columns
            # don't end up with negative content width
            ('VALIGN',        (0, 0), (-1, -1), 'MIDDLE'),
            ('LEFTPADDING',   (0, 1), (-1, -1), 2),
            ('RIGHTPADDING',  (0, 1), (-1, -1), 2),
            ('TOPPADDING',    (0, 2), (-1, -1), 2),
            ('BOTTOMPADDING', (0, 2), (-1, -1), 2),
            ('LINEBELOW',     (0, 1), (-1, -2), 0.25, colors.HexColor('#DDDDDD')),
        ]
        # If empty, span the placeholder across all columns
        if not rows:
            style_cmds.append(('SPAN', (0, 2), (-1, 2)))
        tbl.setStyle(TableStyle(style_cmds))
        return tbl

    def _add_changes_section(self):
        s = self.styles
        self.story.append(SetSource(SOURCE_DEFAULT))
        self.story.append(Paragraph(
            'This page shows the recommendation changes compared to the previous edition of the '
            'EM Bond List, while the reasons for changes include valuation, technical factors, '
            'and/or fundamentals.',
            s['body']))
        self.story.append(Spacer(1, 3*mm))
        self.story.append(Paragraph(
           '<b>Changes to "attractive"</b>: Drawing on our fundamental analysis of the issuer, as presented in our latest CIO report, we believe this specific bond currently trades at a relatively higher yield than it should compared to others with similar risk profiles, so that we now anticipate it will deliver a higher total return.',
           s['body']))
        self.story.append(Spacer(1, 3*mm))
        self.story.append(Paragraph(
           '<b>Changes to "expensive"</b>: Drawing on our fundamental analysis of the issuer, as presented in our latest CIO report, we believe this specific bond currently trades at a relatively lower yield than it should compared to others with similar risk profiles, so that we now anticipate it will deliver a lower total return.',
           s['body']))
        self.story.append(Spacer(1, 3*mm))
        self.story.append(Paragraph(
           '<b>Changes to "fair"</b>: Drawing on our fundamental analysis of the issuer, as presented in our latest CIO report, we believe this specific bond currently trades at an adequate yield compared to others with similar risk profiles, so that we now anticipate it will deliver a similar total return.',
           s['body']))
            
        self.story.append(Spacer(1, 4*mm))

        changes = self.data.recommendation_changes()

        def pair(left_title, left_rows, left_arrow, right_title, right_rows, right_arrow):
            left  = self._build_change_table(left_title,  left_rows,  left_arrow)
            right = self._build_change_table(right_title, right_rows, right_arrow)
            grid = Table(
                [[left, right]],
                colWidths=[CONTENT_WIDTH / 2, CONTENT_WIDTH / 2])
            grid.setStyle(TableStyle([
                ('VALIGN',       (0, 0), (-1, -1), 'TOP'),
                ('LEFTPADDING',  (0, 0), (-1, -1), 0),
                ('RIGHTPADDING', (0, 0), (-1, -1), 4*mm),
                ('TOPPADDING',   (0, 0), (-1, -1), 0),
                ('BOTTOMPADDING',(0, 0), (-1, -1), 0),
            ]))
            return grid

        # Row 1: Upgrades (↗) | Downgrades (↘) — always 2-up.
        self.story.append(pair(
            'Upgrades',   changes['upgrades'],   '\u25B2',
            'Downgrades', changes['downgrades'], '\u25BC',
        ))
        self.story.append(Spacer(1, 5*mm))  

    def _add_additions_deletions_section(self):
       """Standalone Additions & Deletions section — separate page, its own
       title, registered in the TOC via _begin_section."""
       s = self.styles
       self.story.append(SetSource(SOURCE_DEFAULT))
       self.story.append(Paragraph(
           'This page shows the additions and deletions compared to the '
           'previous EM Bond List, while the reasons include newly '
           'initiated coverage or issued bonds for additional and short '
           'time to maturity, or technical factors for deletions.',
           s['body']))
       self.story.append(Spacer(1, 3*mm))
       changes = self.data.recommendation_changes()
       adds = changes['additions']
       dels = changes['deletions']
       # Filter out any rows with missing or blank issuer
       adds = [r for r in adds if (r.get('issuer') or '').strip()]
       dels = [r for r in dels if (r.get('issuer') or '').strip()]
       
       def pair(left_title, left_rows, left_arrow, right_title, right_rows, right_arrow):
           left  = self._build_change_table(left_title,  left_rows,  left_arrow)
           right = self._build_change_table(right_title, right_rows, right_arrow)
           grid = Table(
               [[left, right]],
               colWidths=[CONTENT_WIDTH / 2, CONTENT_WIDTH / 2])
           grid.setStyle(TableStyle([
               ('VALIGN',       (0, 0), (-1, -1), 'TOP'),
               ('LEFTPADDING',  (0, 0), (-1, -1), 0),
               ('RIGHTPADDING', (0, 0), (-1, -1), 4*mm),
               ('TOPPADDING',   (0, 0), (-1, -1), 0),
               ('BOTTOMPADDING',(0, 0), (-1, -1), 0),
           ]))
           return grid
       if not adds and not dels:
           self.story.append(Paragraph('<i>None this week.</i>', s['body']))
       else:
           WIDE_FALLBACK_THRESHOLD = 30
           if max(len(adds), len(dels)) > WIDE_FALLBACK_THRESHOLD:
               self.story.append(self._build_change_table(
                   'Additions', adds, '+', full_width=True))
               self.story.append(Spacer(1, 5*mm))
               self.story.append(self._build_change_table(
                   'Deletions', dels, '\u2212', full_width=True))
           else:
               self.story.append(pair(
                   'Additions', adds, '+',
                   'Deletions', dels, '\u2212',
               ))                               

    # ------------------------------------------------------------- REFERENCE LISTS

    def _add_reference_lists(self):
        """The main reference: bonds grouped by Currency × Region × Grade.

        Produces sections in this order:
            Bonds in USD, Asia
            Bonds in USD, EMEA
            Bonds in USD, GCC
            Bonds in USD, Latin America
            Bonds in EUR, Asia
            ... etc.
        Each section is a separate page (PageBreak before each new currency+region).
        Inside a section: Investment grade first, then Speculative grade; within each,
        SOV before FIN before CORP, alphabetical by country then issuer.
        """
        s = self.styles

        rows = [self.data.bond_row(b) for b in self.data.reference_list_bonds()]

        # Group
        by_ccy_region = defaultdict(list)
        for r in rows:
            by_ccy_region[(r['ccy'], r['region'])].append(r)

        # Preferred currency order (rest alphabetical)
        CCY_ORDER = ['USD', 'EUR', 'GBP', 'CHF', 'CNY', 'SGD', 'MXN', 'BRL',
                     'HKD', 'AUD', 'ZAR', 'IDR', 'INR', 'TRY', 'PLN', 'CAD',
                     'JPY', 'NOK', 'SEK', 'DKK']
        def ccy_sort_key(c):
            try: return (0, CCY_ORDER.index(c))
            except ValueError: return (1, c)

        REGION_ORDER = ['Asia', 'EMEA', 'GCC', 'Latin America']
        def region_sort_key(r):
            try: return REGION_ORDER.index(r)
            except ValueError: return 99

        ccys = sorted({c for c, _ in by_ccy_region.keys()}, key=ccy_sort_key)

        first = True
        for ccy in ccys:
            regions = sorted({r for c, r in by_ccy_region.keys() if c == ccy}, key=region_sort_key)
            for region in regions:
                section_rows = by_ccy_region[(ccy, region)]
                if not section_rows:
                    continue

                section_title = f'Bonds in {currency_name(ccy)}, {region}'
                bookmark_key = f'ref_{ccy}_{region}'.replace(' ', '_')
                # Every reference-list sub-section starts on a fresh page with
                # its title set BEFORE the break (no flicker).
                self._begin_section(bookmark_key, section_title, section_title,
                                    indent=1, page_break=True)
                # On the very first iteration, also plant the top-level
                # 'reflists' bookmark + TOC row on this same page so the TOC
                # 'Reference lists' entry links here.
                if first:
                    self.story.append(BookmarkAnchor('reflists'))
                    if not getattr(self, '_toc_seeded', False):
                        # Insert the top-level row ABOVE the current sub-entry
                        self.toc_sections.insert(len(self.toc_sections) - 1,
                                                 ('reflists', 'Reference lists', 0))
                first = False

                # Split IG / SG
                ig_rows = [r for r in section_rows if r['grade'].startswith('Investment')]
                sg_rows = [r for r in section_rows if r['grade'].startswith('Speculative')]

                if ig_rows:
                   self.story.append(Paragraph(
                       f'<b>Investment grade issuers</b>', s['h3']))
                   self._render_bond_table(ig_rows,
                       show_region_grade_header=False,
                       show_issuer_headers=True)
                if sg_rows:
                   self.story.append(Spacer(1, 3*mm))
                   # Speculative legend
                   self.story.append(Paragraph(
                       f'<b>Speculative grade issuers</b>', s['h3']))
                   self.story.append(Paragraph(
                       f'<i>{SPECULATIVE_LEGEND}</i>', s['spec_warn']))
                   self._render_bond_table(sg_rows,
                       show_region_grade_header=False,
                       show_issuer_headers=True)

    # ------------------------------------------------------------- ISSUER DESCRIPTIONS
    #
    # Two-column layout per issuer (matches the published PDF format):
    #   Left (62 mm):  Issuer name, country, S&P / Moody's rating
    #   Right (rest):  Description paragraph + 8-column credit-view grid
    #                  (0-2 / 2-5 / 5-10 / >10 year maturity buckets plus
    #                   Sub. and Perp. boxes)

    # Credit-view box colours (closer match to the published olive/amber/red)
    _ID_BOX_COLOURS = {
        'green':  colors.HexColor('#6B8E23'),   # olive green
        'yellow': colors.HexColor('#D4A017'),   # amber / gold
        'red':    colors.HexColor('#CC3333'),   # red
    }
    _ID_BOX_EMPTY = colors.HexColor('#E0E0E0')  # light gray when no coverage

    @staticmethod
    def _id_sanitize(text):
        """Escape XML special chars + strip replacement chars for reportlab Paragraph."""
        if not text:
            return ''
        return (text.replace('&', '&amp;').replace('<', '&lt;')
                    .replace('>', '&gt;').replace('\ufffd', ''))

    def _id_get_desc(self, gk):
        """Resolve best-available issuer description (IssuerTexts → issuer update)."""
        texts = self.data.issuer_texts.get(gk, [])
        if texts:
            best_tk, best_desc = '', ''
            for t in texts:
                d = (t.get('IssuerDescription') or '').strip()
                if d and len(d) > 20:
                    tk = t.get('TK', '')
                    if tk > best_tk:
                        best_tk, best_desc = tk, d
            if best_desc:
                return best_desc
        info = self.data.issuer_updates.get(gk, {})
        for field in ('Issuer_Description', 'issuer_comment', 'Issuer_Investment_Case'):
            val = (info.get(field) or '').strip()
            if val and len(val) > 20:
                return val
        return ''

    def _id_senior_colors(self, gk):
        """Return (short, mid, long, very-long) colours for senior/secured cover types."""
        flags = self.data.color_flags.get(gk, {})
        preferred = ('SEN', 'GGB', 'GGL', 'GGA', 'OPF', 'SEC', 'MIX', 'SFF', 'PSN')
        for ct in preferred:
            if ct in flags:
                r = flags[ct]
                sc = (
                    (r.get('WMR Color Short Term Bonds') or '').strip(),
                    (r.get('WMR Color Mid Term Bonds') or '').strip(),
                    (r.get('WMR Color Long Term Bonds') or '').strip(),
                    (r.get('WMR Color Very Long Term Bonds') or '').strip(),
                )
                if any(sc):
                    return sc
        # Fallback: any non-sub/perp cover type that has colours
        for ct, r in flags.items():
            if ct in ('SUB', 'PER', 'CCN', 'HYP', ''):
                continue
            sc = (
                (r.get('WMR Color Short Term Bonds') or '').strip(),
                (r.get('WMR Color Mid Term Bonds') or '').strip(),
                (r.get('WMR Color Long Term Bonds') or '').strip(),
                (r.get('WMR Color Very Long Term Bonds') or '').strip(),
            )
            if any(sc):
                return sc
        return ('', '', '', '')

    def _id_box_bg(self, colour_str):
        if not colour_str:
            return self._ID_BOX_EMPTY
        return self._ID_BOX_COLOURS.get(colour_str.strip().lower(), self._ID_BOX_EMPTY)

    def _id_make_color_grid(self, gk):
        """Build the 8-column credit-view grid for one issuer (matches v2/published)."""
        s = self.styles
        short_c, mid_c, long_c, vlong_c = self._id_senior_colors(gk)
        flags = self.data.color_flags.get(gk, {})
        sub_c = (flags.get('SUB', {}).get('WMR Color Cover Type Level') or '').strip()
        per_c = (flags.get('PER', {}).get('WMR Color Cover Type Level') or '').strip()

        # 8-column layout: label | 4 maturity buckets | spacer-label | Sub | Perp
        gcols = [46*mm, 16*mm, 16*mm, 16*mm, 16*mm, 48*mm, 16*mm, 16*mm]

        grid_hdr = ParagraphStyle('IDGridHdr', parent=s['body_sm'],
                                   fontSize=5.5, leading=7,
                                   textColor=colors.HexColor('#333333'),
                                   alignment=TA_CENTER)
        grid_lbl = ParagraphStyle('IDGridLbl', parent=s['body_sm'],
                                   fontSize=5.5, leading=7,
                                   textColor=colors.HexColor('#333333'))
        na_sty   = ParagraphStyle('IDNa', parent=s['body_sm'],
                                   fontSize=5.5, leading=7,
                                   textColor=colors.HexColor('#666666'),
                                   alignment=TA_CENTER)

        row0 = [
            '',
            Paragraph('0-2Yrs',  grid_hdr),
            Paragraph('2-5Yrs',  grid_hdr),
            Paragraph('5-10Yrs', grid_hdr),
            Paragraph('>10Yrs',  grid_hdr),
            '',
            Paragraph('Sub.',    grid_hdr),
            Paragraph('Perp.',   grid_hdr),
        ]
        sub_cell = '' if sub_c else Paragraph('n.a.', na_sty)
        per_cell = '' if per_c else Paragraph('n.a.', na_sty)
        row1 = [
            Paragraph('UBS credit view on senior bonds:', grid_lbl),
            '', '', '', '',
            Paragraph('UBS credit view on other cover types:', grid_lbl),
            sub_cell,
            per_cell,
        ]

        grid = Table([row0, row1], colWidths=gcols, rowHeights=[10, 13])
        gs = [
            ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
            ('TOPPADDING',    (0, 0), (-1, -1), 0),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 0),
            ('LEFTPADDING',   (0, 0), (-1, -1), 1),
            ('RIGHTPADDING',  (0, 0), (-1, -1), 1),
            ('ALIGN', (1, 0), (4, 0), 'CENTER'),
            ('ALIGN', (6, 0), (7, 0), 'CENTER'),
            ('ALIGN', (6, 1), (7, 1), 'CENTER'),
        ]
        for idx, c in enumerate([short_c, mid_c, long_c, vlong_c]):
            gs.append(('BACKGROUND', (idx + 1, 1), (idx + 1, 1), self._id_box_bg(c)))
        if sub_c: gs.append(('BACKGROUND', (6, 1), (6, 1), self._id_box_bg(sub_c)))
        if per_c: gs.append(('BACKGROUND', (7, 1), (7, 1), self._id_box_bg(per_c)))
        grid.setStyle(TableStyle(gs))
        return grid

    def _add_issuer_descriptions(self):
        """Render the Issuer descriptions section in the published two-column format."""
        s = self.styles
        self.story.append(SetSource(SOURCE_DEFAULT))

        # Per-issuer paragraph styles (kept local so they don't pollute the global set)
        name_sty = ParagraphStyle('IDName', parent=s['body_sm'],
                                   fontSize=8, fontName=self.fonts['bold'], leading=10, spaceAfter=1)
        country_sty = ParagraphStyle('IDCountry', parent=s['body_sm'],
                                      fontSize=7, leading=9,
                                      textColor=colors.HexColor('#333333'))
        rating_sty = ParagraphStyle('IDRating', parent=s['body_sm'],
                                     fontSize=6.5, leading=8,
                                     textColor=colors.HexColor('#666666'))
        comment_sty = ParagraphStyle('IDComment', parent=s['body_sm'],
                                      fontSize=6.5, leading=8)

        # Collect EM issuers referenced by any list bond
        em_gks = set()
        for b in self.data.em_bonds:
            gk = (b.get('GK_Nummer') or '').strip()
            if gk:
                em_gks.add(gk)

        issuer_list = []
        for gk in em_gks:
            name = self.data.issuer_display_name(gk)
            if name:
                issuer_list.append((name, gk))
        issuer_list.sort(key=lambda x: x[0].upper())

        if not issuer_list:
            self.story.append(Paragraph('No issuer descriptions available.', s['body']))
            return

        # Column widths
        left_w    = 62 * mm
        outlook_w = 25 * mm
        right_w   = CONTENT_WIDTH - left_w - outlook_w

        # Header row — repeats on every page via repeatRows=1
        hdr_left = Paragraph(
            '<b>Issuer</b><br/>'
            '<font size="5.5" color="#CCCCCC"><i>Country<br/>'
            'Rating: S&amp;P / Moody\'s</i></font>',
            ParagraphStyle('IDHdrL', parent=s['body_sm'],
                           fontSize=7.5, fontName=self.fonts['bold'],
                           textColor=colors.white, leading=9)
        )
        hdr_outlook = Paragraph(
           '<b>CIO outlook</b>',
           ParagraphStyle('IDHdrO', parent=s['body_sm'],
                          fontSize=7.5, fontName=self.fonts['bold'],
                          textColor=colors.white, leading=9, alignment=TA_CENTER)
       )
        hdr_right = Paragraph(
           '<b>Issuer Comment</b>',
           ParagraphStyle('IDHdrR', parent=s['body_sm'],
                          fontSize=7.5, fontName=self.fonts['bold'],
                          textColor=colors.white, leading=9)
       )
        table_data = [[hdr_left, hdr_outlook, hdr_right]]

        for name, gk in issuer_list:
            cc = self.data.issuer_country_code(gk)
            cn = country_name(cc)

            # S&P / Moody's rating (from issuer-level ratings file)
            ratings = self.data.issuer_ratings.get(gk, {})
            sp  = (ratings.get('SP')  or '').strip()
            mdy = (ratings.get('MDY') or '').strip()
            if sp and mdy:   rat_str = f'{sp} / {mdy}'
            elif sp:         rat_str = sp
            elif mdy:        rat_str = mdy
            else:            rat_str = ''

            # Left cell
            left_parts = [Paragraph(f'<b>{self._id_sanitize(name)}</b>', name_sty)]
            if cn:
                left_parts.append(Paragraph(cn, country_sty))
            if rat_str:
                left_parts.append(Spacer(1, 3*mm))
                left_parts.append(Paragraph(self._id_sanitize(rat_str), rating_sty))

            # Right cell — description + credit-view grid
            right_parts = []
            desc = self._id_get_desc(gk)
            if desc:
                right_parts.append(Paragraph(self._id_sanitize(desc), comment_sty))
            right_parts.append(Spacer(1, 2*mm))
            right_parts.append(self._id_make_color_grid(gk))

            outlook = self.data.issuer_trend(gk)
            outlook_cell = Paragraph(
               outlook,
               ParagraphStyle('IDOutlook', parent=s['body_sm'],
                              fontSize=7, leading=9, alignment=TA_CENTER)
           )
            table_data.append([left_parts, outlook_cell, right_parts])

        main_table = Table(table_data, colWidths=[left_w, outlook_w, right_w], repeatRows=1)
        scmds = [
            ('BACKGROUND', (0, 0), (-1, 0), UBS_HEADER_BG),
            ('VALIGN',        (0, 0), (-1, -1), 'TOP'),
            ('TOPPADDING',    (0, 1), (-1, -1), 4),
            ('BOTTOMPADDING', (0, 1), (-1, -1), 5),
            ('LEFTPADDING',   (0, 0), (-1, -1), 4),
            ('RIGHTPADDING',  (0, 0), (-1, -1), 4),
            ('TOPPADDING',    (0, 0), (-1, 0), 5),
            ('BOTTOMPADDING', (0, 0), (-1, 0), 5),
            ('LINEBELOW',     (0, 0), (-1, 0), 0.5, UBS_DARK),
        ]
        for i in range(1, len(table_data)):
            bg = colors.HexColor('#F8F8F4') if i % 2 == 1 else colors.white
            scmds.append(('BACKGROUND', (0, i), (-1, i), bg))
            scmds.append(('LINEBELOW',  (0, i), (-1, i), 0.25, colors.HexColor('#DDDDDD')))
        main_table.setStyle(TableStyle(scmds))
        self.story.append(main_table)

    # ------------------------------------------------------------- APPENDIX

    # ------------------------------------------------------------- RATING DEFINITIONS REFERENCE

    # Data for the reference page. Each tuple is (S&P, Moody's, Fitch).
    # Spans are expressed as (start_row, span_count, definition_text) —
    # definition appears once, centered across the spanned rows.
    _RATING_DEF_IG_ROWS = [
        ('AAA',  'Aaa',  'AAA'),
        ('AA+',  'Aa1',  'AA+'),
        ('AA',   'Aa2',  'AA'),
        ('AA-',  'Aa3',  'AA-'),
        ('A+',   'A1',   'A+'),
        ('A',    'A2',   'A'),
        ('A-',   'A3',   'A-'),
        ('BBB+', 'Baa1', 'BBB+'),
        ('BBB',  'Baa2', 'BBB'),
        ('BBB-', 'Baa3', 'BBB-'),
    ]
    _RATING_DEF_HY_ROWS = [
        ('BB+',  'Ba1',  'BB+'),
        ('BB',   'Ba2',  'BB'),
        ('BB-',  'Ba3',  'BB-'),
        ('B+',   'B1',   'B+'),
        ('B',    'B2',   'B'),
        ('B-',   'B3',   'B-'),
        ('CCC+', 'Caa1', ''),
        ('CCC',  'Caa2', 'CCC'),
        ('CCC-', 'Caa3', ''),
        ('CC',   'Ca',   'CC'),
        ('C',    'C',    'C'),
        ('D',    'C',    'D'),
    ]
    # (start_index_within_bucket, span, definition) — indices are 0-based
    # within each bucket (IG or HY).
    _RATING_DEF_IG_SPANS = [
        (0, 1, 'Issuer / Bonds have exceptionally strong credit quality. '
               'AAA is the best credit quality.'),
        (1, 3, 'Issuer / Bonds have very strong credit quality.'),
        (4, 3, 'Issuer / Bonds have high credit quality.'),
        (7, 3, 'Issuer / Bonds have adequate credit quality. '
               'This is the lowest Investment Grade category.'),
    ]
    _RATING_DEF_HY_SPANS = [
        (0, 3, 'Issuer / Bonds have weak credit quality. '
               'This is the highest Speculative Grade category.'),
        (3, 3, 'Issuer / Bonds have very weak credit quality.'),
        (6, 3, 'Issuer / Bonds have extremely weak credit quality.'),
        (9, 2, 'Issuer / Bonds have very high risk of default.'),
        (11, 1, 'Obligor failed to make payment on one or more of its '
                'financial commitments. This is the lowest quality of the '
                'Speculative Grade category.'),
    ]

    # Band colors (soft, institutional — not saturated)
    _IG_BAND_COLOR = colors.HexColor('#A7C4A0')   # muted green
    _HY_BAND_COLOR = colors.HexColor('#D9BB92')   # warm tan

    def _add_rating_definitions(self):
        s = self.styles
        self.story.append(SetSource(''))

        # Full-width light-gray header bar
        bar = Table([[Paragraph('<b>Issuer / Bond rating definitions</b>',
                                ParagraphStyle('rd_bar', parent=s['h3'],
                                                fontSize=9, leading=11,
                                                textColor=UBS_DARK))]],
                    colWidths=[CONTENT_WIDTH])
        bar.setStyle(TableStyle([
            ('BACKGROUND',    (0, 0), (-1, -1), UBS_LIGHT),
            ('LEFTPADDING',   (0, 0), (-1, -1), 4),
            ('TOPPADDING',    (0, 0), (-1, -1), 4),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 4),
        ]))
        self.story.append(bar)
        self.story.append(Spacer(1, 1*mm))

        # Column widths. Landscape CONTENT_WIDTH ≈ 277mm.
        band_w = 10*mm
        col_sp = 24*mm
        col_md = 24*mm
        col_ft = 24*mm
        col_def = CONTENT_WIDTH - (band_w + col_sp + col_md + col_ft)

        # Styles
        hdr_style = ParagraphStyle('rd_hdr', parent=s['h3'],
                                    fontSize=8.5, leading=10,
                                    textColor=UBS_DARK, alignment=TA_LEFT)
        hdr_style_c = ParagraphStyle('rd_hdr_c', parent=hdr_style,
                                      alignment=TA_CENTER)
        cell_style = ParagraphStyle('rd_cell', parent=s['body_sm'],
                                     fontSize=8.5, leading=11,
                                     textColor=UBS_DARK)
        cell_style_c = ParagraphStyle('rd_cell_c', parent=cell_style,
                                       alignment=TA_CENTER,
                                       fontName=self.fonts['bold'])
        def_style = ParagraphStyle('rd_def', parent=cell_style,
                                    fontSize=8.5, leading=11)

        # Build data rows
        data = []
        # Header row
        data.append([
            '',
            Paragraph('S&amp;P',      hdr_style_c),
            Paragraph("Moody\u2019s", hdr_style_c),
            Paragraph('Fitch',        hdr_style_c),
            Paragraph('Definition',   hdr_style),
        ])

        # IG rows
        ig_start = 1  # row index in the table
        for sp, md, ft in self._RATING_DEF_IG_ROWS:
            data.append([
                '',  # band column — cell content set via SPAN later
                Paragraph(sp, cell_style_c),
                Paragraph(md, cell_style_c),
                Paragraph(ft, cell_style_c),
                '',  # definition — filled in via spans
            ])

        # HY rows
        hy_start = ig_start + len(self._RATING_DEF_IG_ROWS)
        for sp, md, ft in self._RATING_DEF_HY_ROWS:
            data.append([
                '',
                Paragraph(sp, cell_style_c),
                Paragraph(md, cell_style_c),
                Paragraph(ft, cell_style_c),
                '',
            ])

        # Fill definitions onto the first row of each span; the SPAN command
        # merges them across multiple rows.
        span_cmds = []
        for (off, span, text) in self._RATING_DEF_IG_SPANS:
            r = ig_start + off
            data[r][4] = Paragraph(text, def_style)
            if span > 1:
                span_cmds.append(('SPAN', (4, r), (4, r + span - 1)))
        for (off, span, text) in self._RATING_DEF_HY_SPANS:
            r = hy_start + off
            data[r][4] = Paragraph(text, def_style)
            if span > 1:
                span_cmds.append(('SPAN', (4, r), (4, r + span - 1)))

        # Band text — put the RotatedText into the first row of each bucket
        # (i.e. row index ig_start and hy_start), and SPAN the band column
        # across all rows in that bucket.
        ig_row_count = len(self._RATING_DEF_IG_ROWS)
        hy_row_count = len(self._RATING_DEF_HY_ROWS)
        # The rotated-text flowable needs a height to live in; estimate each
        # data row at ~6.5mm (body_sm at 8.5pt with padding) — close enough.
        est_row_h = 6.5 * mm
        data[ig_start][0] = RotatedText('Investment Grade',
            font_name=self.fonts['bold'], font_size=10,
            color=UBS_DARK, width=band_w, height=ig_row_count * est_row_h)
        data[hy_start][0] = RotatedText('Non-Investment Grade',
            font_name=self.fonts['bold'], font_size=9,
            color=UBS_DARK, width=band_w, height=hy_row_count * est_row_h)

        col_widths = [band_w, col_sp, col_md, col_ft, col_def]
        tbl = Table(data, colWidths=col_widths, repeatRows=1)

        # Style commands
        style = [
            # Header row
            ('BACKGROUND',    (0, 0), (-1, 0), colors.white),
            ('LINEBELOW',     (1, 0), (-1, 0), 0.5, UBS_HEADER_BG),
            ('TOPPADDING',    (0, 0), (-1, 0), 4),
            ('BOTTOMPADDING', (0, 0), (-1, 0), 4),
            ('LEFTPADDING',   (0, 0), (-1, 0), 3),
            ('RIGHTPADDING',  (0, 0), (-1, 0), 3),

            # Overall
            ('VALIGN',        (0, 0), (-1, -1), 'MIDDLE'),
            ('LEFTPADDING',   (1, 1), (-1, -1), 4),
            ('RIGHTPADDING',  (1, 1), (-1, -1), 4),
            ('TOPPADDING',    (1, 1), (-1, -1), 2),
            ('BOTTOMPADDING', (1, 1), (-1, -1), 2),

            # Band column: span across the IG rows and HY rows,
            # fill color, center content
            ('SPAN', (0, ig_start), (0, ig_start + ig_row_count - 1)),
            ('SPAN', (0, hy_start), (0, hy_start + hy_row_count - 1)),
            ('BACKGROUND', (0, ig_start), (0, ig_start + ig_row_count - 1),
             self._IG_BAND_COLOR),
            ('BACKGROUND', (0, hy_start), (0, hy_start + hy_row_count - 1),
             self._HY_BAND_COLOR),
            ('VALIGN',   (0, ig_start), (0, hy_start + hy_row_count - 1), 'MIDDLE'),
            ('ALIGN',    (0, ig_start), (0, hy_start + hy_row_count - 1), 'CENTER'),
            ('LEFTPADDING',   (0, 0), (0, -1), 0),
            ('RIGHTPADDING',  (0, 0), (0, -1), 0),
            ('TOPPADDING',    (0, 0), (0, -1), 0),
            ('BOTTOMPADDING', (0, 0), (0, -1), 0),

            # Subtle horizontal rules between each data row
            ('LINEBELOW', (1, 1), (-1, -2), 0.25, colors.HexColor('#E0E0E0')),

            # Thin gray line separating IG from HY
            ('LINEBELOW', (1, hy_start - 1), (-1, hy_start - 1),
             0.5, UBS_HEADER_BG),

            # Cell vertical alignment
            ('VALIGN', (1, 1), (3, -1), 'MIDDLE'),
            ('VALIGN', (4, 1), (4, -1), 'MIDDLE'),
        ]
        style.extend(span_cmds)
        tbl.setStyle(TableStyle(style))
        self.story.append(tbl)

    # ------------------------------------------------------------- SANCTIONS NOTICE

    _SANCTIONS_PARAGRAPHS = [
        'Please note that all transactions conducted by UBS are consistent with '
        'sanctions regulations imposed by Switzerland, the European Union, the '
        'United Nations, the United Kingdom and the United States, per UBS '
        'global sanctions policy.',

        'US persons are prohibited from purchasing securities of certain '
        'companies designated as being associated with the Chinese Military in '
        'accordance with the amended US Presidential Executive Order 13959 '
        '(dated 3 June 2021).',

        'Under US Sanctions, as part of OFAC new investment prohibition, US '
        'persons are prohibited from purchasing debt and equity securities '
        'issued by an entity in Russia.',

        'In addition, US financial institutions are prohibited from '
        'participation in the secondary market for ruble or non-ruble '
        'denominated bonds issued after March 1, 2022 by the Central Bank of '
        'Russia, the National Wealth Fund of Russia, or the Ministry of '
        'Finance of the Russian Federation.',

        'Under EU and Swiss sanctions, all trading activity (including '
        'divestments, sales, or conversions) of Russian issued securities '
        'that directly or indirectly involve the NSD, is prohibited.',
    ]

    def _add_sanctions_notice(self):
        s = self.styles
        # Source: empty for this page (no agency or pricing source applies).
        self.story.append(SetSource(''))
        body_style = ParagraphStyle(
            'sanctions_body', parent=s['body'],
            fontSize=10, leading=14, textColor=UBS_DARK,
            spaceAfter=8,
        )
        for para in self._SANCTIONS_PARAGRAPHS:
            self.story.append(Paragraph(para, body_style))

    # ------------------------------------------------------------- APPENDIX

    def _add_appendix(self):
        s = self.styles
        self.story.append(SetSource(SOURCE_RATINGS))
        self.story.append(Paragraph(
            'Each research analyst primarily responsible for the content of this research report, '
            'in whole or in part, certifies that with respect to each security or issuer that the analyst '
            'covered in this report: (1) all of the views expressed accurately reflect his or her personal '
            'views about those securities or issuers; and (2) no part of his or her compensation was, is, '
            'or will be, directly or indirectly, related to the specific recommendations or views expressed '
            'by that research analyst in the research report.', s['body_sm']))
        self.story.append(Spacer(1, 3*mm))
        self.story.append(Paragraph(
            '<b>Important disclosures.</b> UBS and/or its affiliates may have a position in, and make a market '
            'in, any of the securities discussed. For a complete list of current disclosures relating to '
            'companies that are the subject of this research, please refer to '
            '<a href="https://www.ubs.com/disclosures">ubs.com/disclosures</a>.', s['body_sm']))
        self.story.append(Spacer(1, 5*mm))
        self.story.append(Paragraph('<i>[Sanctions notice — to be inserted verbatim from the most recently '
            'published GEM List PDF. Contact CIO Publishing or Compliance for the canonical block.]</i>',
            s['body_sm']))


# ══════════════════════════════════════════════════════════════════════════════
# 11. MAIN
# ══════════════════════════════════════════════════════════════════════════════

def main():
    ap = argparse.ArgumentParser(description='Build the CIO Emerging Markets Bond List PDF.')
    ap.add_argument('--bond-data',       required=True)
    ap.add_argument('--issuer-data',     required=True)
    ap.add_argument('--bond-update',     required=True)
    ap.add_argument('--issuer-update',   required=True)
    ap.add_argument('--color-flags',     required=True)
    ap.add_argument('--issuer-texts',    required=True)
    ap.add_argument('--issuer-ratings',  required=True)
    ap.add_argument('--prev-bond-data',  default=None)
    ap.add_argument('--prev-bond-update',default=None)
    ap.add_argument('--logo',            default=DEFAULT_LOGO_PATH)
    ap.add_argument('--output',          default='GEM_List.pdf')
    ap.add_argument('--xlsx-offshore', default='outputs/GEM_List_Offshore.xlsx')
    ap.add_argument('--xlsx-onshore', default='outputs/GEM_List_Onshore.xlsx')
    ap.add_argument('--ladder-output', default='outputs/LatAm_Bond_Ladder.pdf')
    args = ap.parse_args()

    paths = {
        'bond_data':       args.bond_data,
        'issuer_data':     args.issuer_data,
        'bond_update':     args.bond_update,
        'issuer_update':   args.issuer_update,
        'color_flags':     args.color_flags,
        'issuer_texts':    args.issuer_texts,
        'issuer_ratings':  args.issuer_ratings,
        'prev_bond_data':  args.prev_bond_data,
        'prev_bond_update':args.prev_bond_update,
    }

    data = GEMData(paths)
    builder = GEMPDFBuilder(data, args.output,
                             logo_path=args.logo if os.path.exists(args.logo or '') else None)
    builder.build()
   
    # Write audit reports next to the PDF output
    _write_audit_reports(data, args.output)
   #NOTE: do not remove - required by run_weekly.py
    return data 

def _write_audit_reports(data, pdf_output_path):
    """Write subordinated_rule_report.csv and ratings_consistency_report.csv
    next to the PDF output. Both are always written — empty bodies are still
    useful as proof-of-run."""
    out_dir = os.path.dirname(os.path.abspath(pdf_output_path)) or '.'

    sub_path = os.path.join(out_dir, 'subordinated_rule_report.csv')
    rat_path = os.path.join(out_dir, 'ratings_consistency_report.csv')

    sub_cols = ['isin', 'gk', 'issuer', 'cover_type', 'fo_type',
                'wmr_flag',
                'sp_raw', 'mdy_raw', 'sp_token', 'mdy_token',
                'worst_tier', 'rating_status', 'decision', 'reason']
    with open(sub_path, 'w', newline='', encoding='utf-8') as f:
        w = csv.DictWriter(f, fieldnames=sub_cols)
        w.writeheader()
        for row in sorted(data.subordinated_rule_report,
                          key=lambda r: (r['decision'], (r['issuer'] or '').lower(), r['isin'])):
            w.writerow(row)
    print(f'[audit] wrote {sub_path} ({len(data.subordinated_rule_report)} rows)')

    rat_cols = ['isin', 'gk', 'issuer',
                'sp_raw', 'mdy_raw', 'sp_token', 'mdy_token',
                'sp_tier', 'mdy_tier', 'worst_tier',
                'computed_grade', 'source_wmr_flag', 'source_grade',
                'discrepancy']
    with open(rat_path, 'w', newline='', encoding='utf-8') as f:
        w = csv.DictWriter(f, fieldnames=rat_cols)
        w.writeheader()
        # Write discrepancies first, then the rest — easier spot-check
        discreps = [r for r in data.ratings_consistency_report if r['discrepancy']]
        rest     = [r for r in data.ratings_consistency_report if not r['discrepancy']]
        for row in discreps + rest:
            w.writerow(row)
    print(f'[audit] wrote {rat_path} ({len(data.ratings_consistency_report)} rows, '
          f'{sum(1 for r in data.ratings_consistency_report if r["discrepancy"])} discrepancies)')


if __name__ == '__main__':
    main()
