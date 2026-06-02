"""GEM Excel builder.

Architecture: surgical XML manipulation on a binary-copied template.

The two published Excel templates contain embedded charts, shapes, and
text boxes that openpyxl drops on a load+save round-trip. To keep those
intact, this module never round-trips the workbook through openpyxl.
Instead it:

  1. Binary-copies the template .xlsx → output (preserves every file in
     the archive: drawings, charts, shapes, images, printer settings).
  2. Reads ONLY the specific sheet XMLs we need to change, modifies them
     in-place using lxml, and writes them back into the zip.
  3. Updates `sharedStrings.xml` with any new text values introduced by
     the BondList data; leaves `styles.xml` and every other ancillary
     file untouched.
  4. Removes `<sheetProtection .../>` from all sheets so compliance can
     edit them.

What we change each weekly build:
  - Cover sheet: one date cell (`REPORT_DATE_FMT`) and the active tab
    (forced to 'Cover' via `_set_active_sheet` so the file always opens there)
  - BondList sheet: rows 9+ replaced with fresh data
  - 'Changes this week' sheet (onshore only): rows 5+ replaced with the
    onshore-filtered week-on-week diff
  - 'Issuer rating history' / 'Disclosures' wipes: AVAILABLE but currently
    DISABLED — see the note in build_onshore_xlsx (the template ships real
    rating-history content we must not destroy).

Everything else — Cover layout, Terminology, Guidance pages, Disclaimers
US, drawings, conditional formatting, column widths, freeze panes — is
identical to the template.


MAINTAINER'S GUIDE — where to change common things
---------------------------------------------------
  • Cover date wording ..................... REPORT_DATE_FMT (top of CONFIG)
  • Which bonds appear (offshore) .......... is_offshore_eligible()
  • Which bonds appear (US onshore) ........ is_onshore_eligible() + the
       ONSHORE_* sets (excluded sovereigns, quasi-sovereign issuers, overrides)
  • The 23 column values for one bond row .. _compute_row()
       (ratings, region, yield-as-n/a-for-FRNs, 'Top EM Bond List' text, etc.)
  • Region/eligibility/rating LOGIC itself . lives in gem_report_builder_v3.py
       (this module imports GEMData, effective_region, country_name) — change
       it there so the PDF and Excel always agree.

NOTE on shared logic: eligibility, rating resolution, region classification
and the week-on-week diff are all owned by gem_report_builder_v3.GEMData. This
module only decides offshore-vs-onshore inclusion and renders rows into the
template. Keep business rules in the PDF module to avoid PDF/Excel drift.
"""

from __future__ import annotations

import datetime
import os
import re
import shutil
import zipfile

from lxml import etree

from gem_report_builder_v3 import country_name, effective_region

def _has_issuer_name(bond, data):
    gk = (bond.get('GK_Nummer') or '').strip()
    issuer = data.issuer_display_name(gk, fallback=bond.get('IssuerName', '')).strip()
    return bool(issuer)


# ════════════════════════════════════════════════════════════════════════════
# CONFIGURATION
# ════════════════════════════════════════════════════════════════════════════

_HERE = os.path.dirname(os.path.abspath(__file__))
DEFAULT_OFFSHORE_TEMPLATE = os.path.join(_HERE, 'templates', 'Offshore_TEMPLATE.xlsx')
DEFAULT_ONSHORE_TEMPLATE  = os.path.join(_HERE, 'templates', 'Onshore_TEMPLATE.xlsx')

_REPORT_DATE = datetime.date.today()
# Cover-sheet date, e.g. 'May 28, 2026' (no leading zero on the day).
REPORT_DATE_FMT = f'{_REPORT_DATE:%b} {_REPORT_DATE.day}, {_REPORT_DATE.year}'

# Cover-sheet date cell location (sheet name, row, col)
OFFSHORE_COVER_DATE_CELL = ('Cover', 1, 13)
ONSHORE_COVER_DATE_CELL  = ('Cover', 1, 13)

# ── US Onshore eligibility ────────────────────────────────────────────────
ONSHORE_SOV_EXCLUDED_COUNTRIES = {
    'AR', 'VE', 'BH', 'EG', 'KE', 'LK', 'NG',
}

ONSHORE_QUASI_SOVEREIGN_ISSUERS = {
    '0000006188',  # Airport Authority Hong Kong
    '0000014484',  # Export-Import Bank of Thailand
    '0000173778',  # Islamic Development Bank
    '0000188609',  # Export Credit Bank of Turkey
    '0000361417',  # Korea Mine Rehabilitation and Mineral Resources
    '0000407737',  # Export-Import Bank of India
    '0000489700',  # African Export-Import Bank
    '0000510024',  # Export-Import Bank of China
    '0000517024',  # The Arab Energy Fund
    '0000633256',  # International Islamic Liquidity Management
    '0000646764',  # Africa Finance Corporation
    '0000721716',  # Export-Import Bank of Korea (KEXIM)
}

ONSHORE_QUASI_SOVEREIGN_NAME_PATTERNS = (
    'export-import bank', 'export import bank', 'eximbank',
    'export credit bank',
    'islamic development bank', 'islamic liquidity',
    'asian development bank', 'inter-american development bank',
    'african development bank', 'caribbean development bank',
    'african export-import', 'africa finance corporation',
    'world bank', 'international finance corporation',
)

ONSHORE_INCLUDE_OVERRIDES: set[str] = set()
ONSHORE_EXCLUDE_OVERRIDES: set[str] = set()


def _is_quasi_sovereign(gk, name):
    if gk in ONSHORE_QUASI_SOVEREIGN_ISSUERS:
        return True
    nm = (name or '').lower()
    return any(pat in nm for pat in ONSHORE_QUASI_SOVEREIGN_NAME_PATTERNS)


# CINS prefixes assigned to non-US issuers. An ISIN beginning with one of
# these is a CINS, not a true US CUSIP — i.e. the bond was placed offshore
# (Reg S) rather than SEC-registered. Used to block Reg S Supra and
# Quasi-Sovereign issuance from the US Onshore list.
CINS_FOREIGN_PREFIXES = {
   'USY', 'USG', 'USN', 'USP', 'USC', 'USU',
   'USL', 'USA', 'USV',
}

def _is_reg_s(bond):
   """True if the bond appears to be Reg S (offshore-only).
   Two complementary signals; either one is sufficient:
     • ISIN is a CINS (non-US issuer prefix) — not a true US CUSIP.
     • CIO_Market_Of_Issuance is 'International' — issuer flagged the
       placement as international/offshore.
   """
   isin = (bond.get('Isin') or '').strip()
   if not isin.startswith('US') or isin[:3] in CINS_FOREIGN_PREFIXES:
       return True
   market = (bond.get('CIO_Market_Of_Issuance') or '').strip()
   if market == 'International':
       return True
   return False

def is_onshore_eligible(bond, data, upd=None):
   """Decide whether a single bond record qualifies for the US Onshore list.

   The same function feeds both the BondList sheet AND the Onshore "Changes
   this week" sheet (via _onshore_change in build_onshore_xlsx), so any
   change here propagates symmetrically to the week-on-week diff — a bond
   that becomes newly eligible shows up as an Addition; one that loses
   eligibility shows up as a Deletion.

   THE RULES, in order:

     1. Bond is denominated in USD.
     2. Bond is tagged as part of the GEM universe (TopListCategory='GEM').
     3. Bond is more than 180 days from maturity (except defaulted
        Venezuelan sovereigns, which we keep until a restructuring).
     4. Per-ISIN overrides win over the rest, in this order:
          • ONSHORE_EXCLUDE_OVERRIDES — force-exclude specific bonds.
          • ONSHORE_INCLUDE_OVERRIDES — force-include specific bonds.
     5. Issuer-type based US-eligibility:
          • SOVEREIGN          — eligible unless the country is on the
                                  ONSHORE_SOV_EXCLUDED_COUNTRIES list.
                                  (Sovereigns are SEC-registered globally
                                  including sovereign Sukuks with CINS
                                  prefixes, so the ISIN-prefix Reg S test
                                  is skipped.)
          • SUPRANATIONAL      — ELIGIBLE (Reg S test skipped). World Bank
                                  / IFC / regional development banks
                                  receive blanket inclusion, matching how
                                  the published list has historically
                                  treated them.
          • QUASI-SOVEREIGN    — ELIGIBLE (Reg S test skipped). Eximbanks,
                                  Islamic Development Bank, Africa Finance
                                  Corporation, etc. — same treatment as
                                  supras.
          • CORPORATE/FINANCIAL — eligible only if the bond is NOT Reg S.
                                  "Reg S" here means: ISIN prefix is CINS
                                  (non-US issuer ID) OR Market of Issuance
                                  is 'International'. This keeps 144A-only
                                  corporate paper off the Onshore list.

   This is the explicit rule set — there is no external dependency. The
   rationale for the SUPRA/QUASI carve-out (versus the corporate rule) is
   documented in the maintainer's guide at the top of this file.
   """
   # 1. Currency
   if (bond.get('CCY') or '').strip().upper() != 'USD':
       return False
   isin = (bond.get('Isin') or '').strip()
   if upd is None:
       upd = data.bond_updates.get(isin, {})
   # 2. GEM universe membership
   if (upd.get('TopListCategory') or '').strip() != 'GEM':
       return False
   # 3. Maturity cutoff — drop bonds maturing within 180 days, except
   # defaulted Venezuelan sovereigns (kept until a restructuring).
   maturity = (bond.get('Maturity') or '').strip()
   if maturity:
       mat_date = None
       for fmt in ('%d.%m.%Y', '%Y-%m-%d', '%d/%m/%Y', '%m/%d/%Y', '%d-%m-%Y'):
           try:
               mat_date = datetime.datetime.strptime(maturity, fmt).date()
               break
           except ValueError:
               continue
       if mat_date and (mat_date - datetime.date.today()).days < 180:
           gk = (bond.get('GK_Nummer') or '').strip()
           is_ve_sov = (data.issuer_type(gk) == 'SOV'
                        and data.issuer_country_code(gk) == 'VE')
           if not is_ve_sov:
               return False
   # 4. Per-ISIN overrides (force-exclude wins over force-include).
   if isin in ONSHORE_EXCLUDE_OVERRIDES:
       return False
   if isin in ONSHORE_INCLUDE_OVERRIDES:
       return True
   # 5. Issuer-type based eligibility.
   gk = (bond.get('GK_Nummer') or '').strip()
   itype = data.issuer_type(gk)
   if itype == 'SOV':
       return data.issuer_country_code(gk) not in ONSHORE_SOV_EXCLUDED_COUNTRIES
   if itype == 'SUPRA':
       # Supranationals: blanket inclusion (no Reg S filter applied).
       # World Bank, IFC, regional development banks, CABEI, etc.
       return True
   name = data.issuer_display_name(gk, fallback=bond.get('IssuerName', ''))
   if _is_quasi_sovereign(gk, name):
       # Quasi-sovereigns (Eximbanks, Islamic Development Bank, Africa
       # Finance Corp, etc.): blanket inclusion (no Reg S filter applied).
       return True
   # Corporates / financials: the Reg S exclusion applies — must be a
   # true US CUSIP and not flagged as International placement.
   if _is_reg_s(bond):
       return False
   return True



def is_offshore_eligible(bond, data):
   isin = (bond.get('Isin') or '').strip()
   upd = data.bond_updates.get(isin, {})
   if (upd.get('TopListCategory') or '').strip() != 'GEM':
       return False
   # Maturity cutoff — drop bonds maturing within 180 days, except
   # defaulted Venezuelan sovereigns (kept until a restructuring).
   maturity = (bond.get('Maturity') or '').strip()
   if maturity:
       mat_date = None
       for fmt in ('%d.%m.%Y', '%Y-%m-%d', '%d/%m/%Y', '%m/%d/%Y', '%d-%m-%Y'):
           try:
               mat_date = datetime.datetime.strptime(maturity, fmt).date()
               break
           except ValueError:
               continue
       if mat_date and (mat_date - datetime.date.today()).days < 180:
           gk = (bond.get('GK_Nummer') or '').strip()
           is_ve_sov = (data.issuer_type(gk) == 'SOV'
                        and data.issuer_country_code(gk) == 'VE')
           if not is_ve_sov:
               return False
   return True

# ════════════════════════════════════════════════════════════════════════════
# DATA HELPERS
# ════════════════════════════════════════════════════════════════════════════

EXCEL_EPOCH = datetime.datetime(1899, 12, 30)


def _date_to_excel_serial(d):
    if isinstance(d, datetime.date) and not isinstance(d, datetime.datetime):
        d = datetime.datetime(d.year, d.month, d.day)
    return (d - EXCEL_EPOCH).days


def _num(v):
    if v is None or v == '':
        return None
    try:
        return float(v)
    except (TypeError, ValueError):
        return None


def _parse_date(v):
    if not v:
        return None
    if isinstance(v, (datetime.date, datetime.datetime)):
        return v
    s = str(v).strip()
    for fmt in (
        '%d.%m.%Y',  # 31.12.2029
        '%Y-%m-%d',  # 2029-12-31
        '%d/%m/%Y',  # 31/12/2029
        
        '%m/%d/%Y',  # 12-31-2029
        '%d-%m-%Y',  # 31/12/2029
        '%Y/%m/%d',  # 2029/12/31
    ):
        try:
            return datetime.datetime.strptime(s, fmt)
        except ValueError:
            continue
    return None


def _col_letter(n):
    """1 → A, 27 → AA"""
    s = ''
    while n > 0:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s
    return s


def _compute_row(bond, data, schema):
    """Build a list of 23 cell values for one bond row."""
    isin   = (bond.get('Isin') or '').strip()
    valor  = (bond.get('Valor') or '').strip()
    gk     = (bond.get('GK_Nummer') or '').strip()
    upd    = data.bond_updates.get(isin, {})

    rec_code = (upd.get('WMR_Bond_Recommendation')
                or bond.get('WMR_Bond_Recommendation') or '').strip().upper()
    view = {'OP': 'attr.', 'UP': 'exp.', 'SELL': 'sell'}.get(rec_code, 'fair')
    issuer = data.issuer_display_name(gk, fallback=bond.get('IssuerName', ''))
    cc = data.issuer_country_code(gk)
    # Pass `upd` so subordinated bonds pick up the bond-level (notched-down)
    # rating from PublishableBondDataUpdate.RatingSP/RatingMdy rather than
    # the higher issuer-level rating.
    eff = data.effective_issuer_rating(gk, bond, upd)
    sp  = eff['sp_token'] or '-'
    mdy = eff['mdy_token'] or '-'
    coupon  = _num(bond.get('Coupon'))
    if coupon is not None:
        coupon = coupon / 100
    px      = _num(bond.get('PXASK_ExecDesk'))
    # YTM is meaningless for floaters — the ExecDesk feed often emits 0.0.
    # Replace with the string 'n/a' so Excel doesn't show a misleading 0%.
    cpn_type_raw = (bond.get('CpnType') or '').strip().lower()
    fo_type_raw  = (bond.get('FOType')  or '').strip().lower()
    is_floater   = (cpn_type_raw in ('variable', 'fixed/variable') or
                    'float' in fo_type_raw)
    yld     = 'n/a' if is_floater else _num(bond.get('YLDASK_ExecDesk'))
    minamt  = _num(bond.get('MinAmt'))
    mininc  = _num(bond.get('MinInc'))
    amtout  = _num(bond.get('AmtOutstanding'))
    mat_raw = (bond.get('Maturity') or '').strip()
    mat_dt = _parse_date(mat_raw)
    maturity = mat_dt.strftime('%m/%d/%Y') if mat_dt else (mat_raw if mat_raw else 'Perpetual')
    # 'Top EM Bond List' column: spell out the full label in each flagged
    # cell (matches the old production file); leave non-top bonds truly blank.
    # Applies to both offshore and onshore — _compute_row feeds both.
    is_top = (bond.get('Product_Use') or '').strip() == '7' and rec_code != 'SELL'
    top_list = 'Top EM Bond List' if is_top else None
    green_raw = (bond.get('GreenBond') or '').strip().upper()
    green_label = {'G': 'green', 'S': 'social', 'U': 'sustainable'}.get(green_raw, '-')
    outlook = data.issuer_trend(gk) if gk else 'Stable'
    rf = (bond.get('WMRColorFlag') or '').strip().upper()
    if rf not in ('GREEN', 'YELLOW', 'RED'):
        rf = '-'
    analyst = data.analyst_for_gk(gk) or '-'
    comment = (bond.get('WMR_Bond_Comment') or '').strip() or '-'

    # Restrictions: authoritative lookup via the shared GEMData logic so the
    # Excel and PDF always agree. Reads the PRIIPS reference keyed by ISIN
    # (Valoren fallback); falls back to the legacy heuristic only when no
    # reference file was supplied. NOTE: this replaces the old behaviour that
    # forced a minimum of "1" on every bond — bonds the reference marks
    # Non-Complex now correctly show blank.
    restrictions_str = data.restriction_for(isin, valor, bond)
    issuer_desc = data.issuer_description_text(gk) or '-'

    row = [
        view, isin, (int(valor) if valor.isdigit() else valor),
        issuer, (bond.get('CCY') or '').strip().upper(),
        coupon, px, yld, maturity,
        sp, mdy,
        minamt, mininc, amtout,
        effective_region(issuer, cc),
        (country_name(cc) if cc else ''),
        top_list, green_label,
        outlook, rf,
        analyst, comment,
    ]
    if schema == 'offshore':
        row.append(restrictions_str)
    else:
        row.append(issuer_desc)
    return row


# ════════════════════════════════════════════════════════════════════════════
# XML HELPERS — surgical sheet manipulation
# ════════════════════════════════════════════════════════════════════════════

NS_MAIN = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
NS_R    = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
NSMAP   = {'m': NS_MAIN, 'r': NS_R}


class _XlsxEditor:
    """In-memory xlsx editor that writes everything back atomically."""

    def __init__(self, path):
        self.path = path
        with zipfile.ZipFile(path, 'r') as z:
            self.files = {n: z.read(n) for n in z.namelist()}
        self._sheet_map = self._build_sheet_map()

    def _build_sheet_map(self):
        wb = etree.fromstring(self.files['xl/workbook.xml'])
        rels = etree.fromstring(self.files['xl/_rels/workbook.xml.rels'])
        rel_target = {r.get('Id'): r.get('Target')
                      for r in rels.findall('.//{*}Relationship')}
        out = {}
        for s in wb.findall('.//m:sheet', NSMAP):
            name = s.get('name')
            rid  = s.get(f'{{{NS_R}}}id')
            target = rel_target.get(rid)
            if target:
                if not target.startswith('/'):
                    target = 'xl/' + target
                else:
                    target = target.lstrip('/')
                out[name] = target
        return out

    def sheet_path(self, name):
        return self._sheet_map.get(name)

    def get_sheet_xml(self, name):
        path = self.sheet_path(name)
        if not path:
            return None, None
        return path, etree.fromstring(self.files[path])

    def set_sheet_xml(self, name, root):
        path = self.sheet_path(name)
        if not path:
            return
        self.files[path] = etree.tostring(root, xml_declaration=True,
                                          encoding='UTF-8', standalone=True)

    # ── Shared strings ─────────────────────────────────────────────────
    def _load_shared_strings(self):
        if not hasattr(self, '_ss_root'):
            if 'xl/sharedStrings.xml' in self.files:
                self._ss_root = etree.fromstring(self.files['xl/sharedStrings.xml'])
            else:
                self._ss_root = etree.fromstring(
                    f'<sst xmlns="{NS_MAIN}" count="0" uniqueCount="0"/>'.encode())
            self._ss_count = 0  # total <si> element count (positional indexing)
            self._ss_index = {}
            for i, si in enumerate(self._ss_root.findall('m:si', NSMAP)):
                self._ss_count = i + 1
                # Concat all <t> sub-elements to handle rich-text <r><t>...</t></r>
                ts = si.findall('.//m:t', NSMAP)
                text = ''.join(t.text or '' for t in ts)
                # First-occurrence wins: cell refs in the template that point
                # to the first instance stay valid.
                if text and text not in self._ss_index:
                    self._ss_index[text] = i

    def add_string(self, s):
        self._load_shared_strings()
        if s in self._ss_index:
            return self._ss_index[s]
        si = etree.SubElement(self._ss_root, f'{{{NS_MAIN}}}si')
        t  = etree.SubElement(si, f'{{{NS_MAIN}}}t')
        if s != s.strip():
            t.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
        t.text = s
        idx = self._ss_count
        self._ss_count += 1
        self._ss_index[s] = idx
        return idx

    def _save_shared_strings(self):
        if hasattr(self, '_ss_root'):
            self._ss_root.set('count', str(self._ss_count))
            self._ss_root.set('uniqueCount', str(self._ss_count))
            self.files['xl/sharedStrings.xml'] = etree.tostring(
                self._ss_root, xml_declaration=True,
                encoding='UTF-8', standalone=True)
            # Make sure [Content_Types].xml lists sharedStrings (it should
            # already from the template, but defensively add it if missing)
            ct_xml = self.files['[Content_Types].xml'].decode()
            if 'sharedStrings.xml' not in ct_xml:
                ct_xml = ct_xml.replace(
                    '</Types>',
                    '<Override PartName="/xl/sharedStrings.xml" '
                    'ContentType="application/vnd.openxmlformats-officedocument.'
                    'spreadsheetml.sharedStrings+xml"/></Types>')
                self.files['[Content_Types].xml'] = ct_xml.encode()

    # ── Save ──────────────────────────────────────────────────────────
    def save(self):
        self._save_shared_strings()
        tmp = self.path + '.tmp'
        with zipfile.ZipFile(tmp, 'w', zipfile.ZIP_DEFLATED) as zout:
            for name, content in self.files.items():
                zout.writestr(name, content)
        shutil.move(tmp, self.path)


# ════════════════════════════════════════════════════════════════════════════
# SHEET-LEVEL OPERATIONS
# ════════════════════════════════════════════════════════════════════════════

def _replace_ubslogo_with_text(editor: _XlsxEditor, replacement='UBS'):
    """OPTIONAL / currently unused. Call sites are commented out in the
    builders below — kept available in case the proprietary 'UBSLogo' font is
    unavailable on a reviewer's machine and the cover logo renders as 'AB'.

    The templates render the UBS logo via an 'a'/'b' pair in the
    proprietary 'UBSLogo' font. On machines without that font (most
    compliance reviewers) this shows as the literal letters 'AB'. Replace
    the whole pair with a single readable 'UBS' text run in red Arial so
    the cover reads correctly everywhere.
    """
    # Match the 2-run 'a' + 'b' pair using typeface="UBSLogo".
    pair_re = re.compile(
        r'<a:r>\s*<a:rPr[^>]*>.*?typeface="UBSLogo".*?</a:rPr>\s*<a:t>a</a:t>\s*</a:r>'
        r'\s*<a:r>\s*<a:rPr[^>]*>.*?typeface="UBSLogo".*?</a:rPr>\s*<a:t>b</a:t>\s*</a:r>',
        re.DOTALL)
    new_run = (
        '<a:r><a:rPr lang="en-US" sz="2400" b="1" i="0" u="none" '
        'strike="noStrike" baseline="0">'
        '<a:solidFill><a:srgbClr val="E60000"/></a:solidFill>'
        '<a:latin typeface="Arial"/></a:rPr>'
        f'<a:t>{replacement}</a:t></a:r>'
    )
    changed = 0
    for name in list(editor.files.keys()):
        if name.startswith('xl/drawings/') and name.endswith('.xml'):
            content = editor.files[name].decode()
            new_content, n = pair_re.subn(new_run, content)
            if n:
                editor.files[name] = new_content.encode()
                changed += n
    return changed


def _unlock_all_sheets(editor: _XlsxEditor):
    for name, path in editor._sheet_map.items():
        root = etree.fromstring(editor.files[path])
        # Remove sheetProtection elements
        for sp in root.findall('m:sheetProtection', NSMAP):
            root.remove(sp)
        editor.files[path] = etree.tostring(
            root, xml_declaration=True, encoding='UTF-8', standalone=True)


def _set_active_sheet(editor: _XlsxEditor, sheet_name='Cover'):
    """Force the workbook to open on `sheet_name` (the Cover by default),
    regardless of which sheet was active when the template was last saved.

    Two things control the on-open view:
      1. <workbookView activeTab="N"/> in xl/workbook.xml — the index of the
         tab Excel selects on open.
      2. <sheetView tabSelected="1"/> on each worksheet — only the target
         sheet should carry it; leaving it on others makes Excel open with a
         multi-sheet selection (and can override activeTab).
    We also reset the Cover's selection to A1 so it opens at the top-left.
    """
    # 1) workbook.xml — point activeTab at the target sheet's position
    wb_root = etree.fromstring(editor.files['xl/workbook.xml'])
    sheets = wb_root.findall('.//m:sheets/m:sheet', NSMAP)
    target_idx = next((i for i, s in enumerate(sheets)
                       if s.get('name') == sheet_name), None)
    if target_idx is None:
        return
    book_views = wb_root.find('m:bookViews', NSMAP)
    if book_views is None:
        # bookViews must precede <sheets> per the schema — insert it there.
        book_views = etree.Element(f'{{{NS_MAIN}}}bookViews')
        wb_root.find('m:sheets', NSMAP).addprevious(book_views)
    wv = book_views.find('m:workbookView', NSMAP)
    if wv is None:
        wv = etree.SubElement(book_views, f'{{{NS_MAIN}}}workbookView')
    wv.set('activeTab', str(target_idx))
    editor.files['xl/workbook.xml'] = etree.tostring(
        wb_root, xml_declaration=True, encoding='UTF-8', standalone=True)

    # 2) per-sheet — tabSelected only on the target; reset its selection to A1
    for name, path in editor._sheet_map.items():
        root = etree.fromstring(editor.files[path])
        svs = root.find('m:sheetViews', NSMAP)
        if svs is None:
            continue
        for sv in svs.findall('m:sheetView', NSMAP):
            if name == sheet_name:
                sv.set('tabSelected', '1')
                sv.set('topLeftCell', 'A1')
                for sel in sv.findall('m:selection', NSMAP):
                    sv.remove(sel)
                sel = etree.SubElement(sv, f'{{{NS_MAIN}}}selection')
                sel.set('activeCell', 'A1')
                sel.set('sqref', 'A1')
            else:
                sv.attrib.pop('tabSelected', None)
        editor.files[path] = etree.tostring(
            root, xml_declaration=True, encoding='UTF-8', standalone=True)


def _set_cell_inline(editor: _XlsxEditor, sheet_name, row_num, col_num, value):
    """Update a single cell's value (preserving its style index) in a sheet.
    Uses inline string for text, raw value for numbers/dates."""
    path, root = editor.get_sheet_xml(sheet_name)
    if root is None:
        return
    sd = root.find('m:sheetData', NSMAP)
    if sd is None:
        return
    cell_ref = f'{_col_letter(col_num)}{row_num}'
    # Find the row
    target_row = None
    for r in sd.findall('m:row', NSMAP):
        if r.get('r') == str(row_num):
            target_row = r
            break
    if target_row is None:
        target_row = etree.SubElement(sd, f'{{{NS_MAIN}}}row')
        target_row.set('r', str(row_num))
    # Find or create the cell
    target_cell = None
    for c in target_row.findall('m:c', NSMAP):
        if c.get('r') == cell_ref:
            target_cell = c
            break
    if target_cell is None:
        target_cell = etree.SubElement(target_row, f'{{{NS_MAIN}}}c')
        target_cell.set('r', cell_ref)
    # Clear children
    for child in list(target_cell):
        target_cell.remove(child)
    # Set value as inline string (most reliable)
    if value is None:
        target_cell.attrib.pop('t', None)
    elif isinstance(value, (int, float)) and not isinstance(value, bool):
        target_cell.attrib.pop('t', None)
        v = etree.SubElement(target_cell, f'{{{NS_MAIN}}}v')
        v.text = str(value)
    else:
        target_cell.set('t', 'inlineStr')
        is_el = etree.SubElement(target_cell, f'{{{NS_MAIN}}}is')
        t = etree.SubElement(is_el, f'{{{NS_MAIN}}}t')
        t.text = str(value)
        if str(value) != str(value).strip():
            t.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
    editor.set_sheet_xml(sheet_name, root)


def _replace_sheet_data_rows(editor: _XlsxEditor, sheet_name,
                              keep_first_n_rows, new_rows_xml,
                              uniform_row_height=14.0):
    """Keep rows 1..keep_first_n_rows from the template; remove all rows
    after that; append new_rows_xml (a single XML string of <row> elements).
    """
    path, root = editor.get_sheet_xml(sheet_name)
    if root is None:
        return
    sd = root.find('m:sheetData', NSMAP)
    if sd is None:
        sd = etree.SubElement(root, f'{{{NS_MAIN}}}sheetData')
    # Remove rows > keep_first_n_rows
    for r in list(sd.findall('m:row', NSMAP)):
        if int(r.get('r', 0)) > keep_first_n_rows:
            sd.remove(r)
    # Parse and append new rows
    if new_rows_xml:
        wrapper = etree.fromstring(
            f'<wrap xmlns="{NS_MAIN}">{new_rows_xml}</wrap>'.encode())
        for r in wrapper.findall('m:row', NSMAP):
            sd.append(r)
    editor.set_sheet_xml(sheet_name, root)


def _wipe_sheet_to_note(editor: _XlsxEditor, sheet_name, note):
    """Replace ALL content in a sheet with a single note in cell A1.
    Removes drawings/refs that referenced specific cells too.

    OPTIONAL / currently unused. The only call sites (in build_onshore_xlsx)
    are commented out, because the current onshore template ships real
    'Issuer rating history' content that must be preserved. Kept available in
    case a future workflow wants those sheets cleared for downstream teams.
    """
    path, root = editor.get_sheet_xml(sheet_name)
    if root is None:
        return
    sd = root.find('m:sheetData', NSMAP)
    if sd is not None:
        for r in list(sd.findall('m:row', NSMAP)):
            sd.remove(r)
    # Remove merged cells references
    for mc in list(root.findall('m:mergeCells', NSMAP)):
        root.remove(mc)
    # Add the note
    sd = root.find('m:sheetData', NSMAP)
    if sd is None:
        sd = etree.SubElement(root, f'{{{NS_MAIN}}}sheetData')
    row = etree.SubElement(sd, f'{{{NS_MAIN}}}row')
    row.set('r', '1')
    cell = etree.SubElement(row, f'{{{NS_MAIN}}}c')
    cell.set('r', 'A1')
    cell.set('t', 'inlineStr')
    is_el = etree.SubElement(cell, f'{{{NS_MAIN}}}is')
    t = etree.SubElement(is_el, f'{{{NS_MAIN}}}t')
    t.text = note
    editor.set_sheet_xml(sheet_name, root)


# ════════════════════════════════════════════════════════════════════════════
# BONDLIST + CHANGES XML BUILDERS
# ════════════════════════════════════════════════════════════════════════════

def _borrow_styles_from_row(editor: _XlsxEditor, sheet_name, row_num):
    """Read style indexes (s="..." attr) from each cell in the given row.
    Used to apply the template's row 9 styling to all our new data rows."""
    path, root = editor.get_sheet_xml(sheet_name)
    if root is None:
        return {}
    sd = root.find('m:sheetData', NSMAP)
    if sd is None:
        return {}
    for r in sd.findall('m:row', NSMAP):
        if r.get('r') == str(row_num):
            styles = {}
            for c in r.findall('m:c', NSMAP):
                ref = c.get('r')
                # Strip row to get column letters
                col = re.match(r'([A-Z]+)\d+', ref).group(1)
                if c.get('s'):
                    styles[col] = c.get('s')
            return styles
    return {}


def _build_bondlist_rows_xml(editor: _XlsxEditor, bond_rows, data, schema,
                              sheet_name='BondList', start_row=9,
                              uniform_height=14.0):
    """Build the XML string for new BondList data rows. Borrows style indexes
    from the template's row `start_row` so the output looks like the template
    intended (borders, fonts, colors, number formats)."""
    styles = _borrow_styles_from_row(editor, sheet_name, start_row)

    # Per-column override: even if template's row 9 happens to use a text
    # style, force a date-format style for the maturity column. Pick the
    # style index from styles.xml if available, else fall back to row 9's.
    # We accept the template's style by default — if dates display as
    # numbers, the template's row 9 style was wrong (column I was '-').
    # In that case we transparently substitute a date-formatted style.
    date_style_idx = _find_or_create_date_style(editor)
    if date_style_idx and 'I' in styles:
        styles['I'] = date_style_idx

    parts = []
    for offset, bond in enumerate(bond_rows):
        row_num = start_row + offset
        values = _compute_row(bond, data, schema)
        cells_xml = []
        for col_idx, val in enumerate(values, start=1):
            col = _col_letter(col_idx)
            ref = f'{col}{row_num}'
            style_attr = f' s="{styles[col]}"' if col in styles else ''
            if val is None:
                cells_xml.append(f'<c r="{ref}"{style_attr}/>')
            elif isinstance(val, datetime.datetime) or isinstance(val, datetime.date):
                serial = _date_to_excel_serial(val)
                cells_xml.append(f'<c r="{ref}"{style_attr}><v>{serial}</v></c>')
            elif isinstance(val, (int, float)) and not isinstance(val, bool):
                cells_xml.append(f'<c r="{ref}"{style_attr}><v>{val}</v></c>')
            else:
                # Use shared string for text (more efficient for repeated values)
                idx = editor.add_string(str(val))
                cells_xml.append(
                    f'<c r="{ref}"{style_attr} t="s"><v>{idx}</v></c>')
        ht_attr = f' ht="{uniform_height}" customHeight="1"'
        parts.append(f'<row r="{row_num}"{ht_attr}>{"".join(cells_xml)}</row>')
    return ''.join(parts)


def _add_nonbold_variant(editor: _XlsxEditor, source_style_idx: int) -> str:
    """Clone an existing cellXfs style but force non-bold font. Returns the
    new style index as a string. Lets us reuse the template's look (fill,
    borders, alignment, number format) while changing just the weight."""
    styles_xml = editor.files['xl/styles.xml']
    root = etree.fromstring(styles_xml)
    cxfs = root.find('m:cellXfs', NSMAP)
    all_xfs = cxfs.findall('m:xf', NSMAP)
    if source_style_idx >= len(all_xfs):
        return str(source_style_idx)
    src_xf = all_xfs[source_style_idx]

    fonts = root.find('m:fonts', NSMAP)
    src_font_id = int(src_xf.get('fontId', 0))
    all_fonts = fonts.findall('m:font', NSMAP)
    if src_font_id >= len(all_fonts):
        return str(source_style_idx)
    src_font = all_fonts[src_font_id]

    # Clone the font, drop any <b/>
    new_font = etree.fromstring(etree.tostring(src_font))
    for b in new_font.findall('m:b', NSMAP):
        new_font.remove(b)
    fonts.append(new_font)
    fonts.set('count', str(len(fonts.findall('m:font', NSMAP))))
    new_font_id = len(fonts.findall('m:font', NSMAP)) - 1

    # Clone the xf, point at the new font
    new_xf = etree.fromstring(etree.tostring(src_xf))
    new_xf.set('fontId', str(new_font_id))
    new_xf.set('applyFont', '1')
    cxfs.append(new_xf)
    cxfs.set('count', str(len(cxfs.findall('m:xf', NSMAP))))
    new_idx = len(cxfs.findall('m:xf', NSMAP)) - 1

    editor.files['xl/styles.xml'] = etree.tostring(
        root, xml_declaration=True, encoding='UTF-8', standalone=True)
    return str(new_idx)


def _find_or_create_date_style(editor: _XlsxEditor):
    """Find a style in styles.xml whose number format is a date, or add one.
    Returns the style index as a string, or None if styles.xml is missing."""
    if 'xl/styles.xml' not in editor.files:
        return None
    styles_xml = editor.files['xl/styles.xml']
    root = etree.fromstring(styles_xml)
    # Built-in date format codes: 14 = m/d/yyyy, 15 = d-mmm-yy, 22 = m/d/yyyy h:mm
    # Custom format we'll add: "dd.mm.yyyy"
    target_fmt_code = 'mm/dd/yyyy'

    nfmts = root.find('m:numFmts', NSMAP)
    if nfmts is None:
        nfmts = etree.Element(f'{{{NS_MAIN}}}numFmts')
        # Insert as first element so it precedes <fonts>
        root.insert(0, nfmts)
    # Check existing
    existing_id = None
    for f in nfmts.findall('m:numFmt', NSMAP):
        if f.get('formatCode') == target_fmt_code:
            existing_id = f.get('numFmtId')
            break
    if existing_id is None:
        # Find next free id (164+ is custom range)
        used_ids = {int(f.get('numFmtId')) for f in nfmts.findall('m:numFmt', NSMAP)}
        new_id = 164
        while new_id in used_ids:
            new_id += 1
        nf = etree.SubElement(nfmts, f'{{{NS_MAIN}}}numFmt')
        nf.set('numFmtId', str(new_id))
        nf.set('formatCode', target_fmt_code)
        nfmts.set('count', str(len(nfmts.findall('m:numFmt', NSMAP))))
        existing_id = str(new_id)

    # Find or create a cellXfs entry that uses this numFmtId
    cxfs = root.find('m:cellXfs', NSMAP)
    if cxfs is None:
        return None
    target_idx = None
    for i, xf in enumerate(cxfs.findall('m:xf', NSMAP)):
        if xf.get('numFmtId') == existing_id:
            target_idx = i
            break
    if target_idx is None:
        # Create a new cellXfs entry — copy attributes from the first xf and override numFmtId
        first_xf = cxfs.find('m:xf', NSMAP)
        new_xf = etree.SubElement(cxfs, f'{{{NS_MAIN}}}xf')
        if first_xf is not None:
            for k, v in first_xf.attrib.items():
                new_xf.set(k, v)
        new_xf.set('numFmtId', existing_id)
        new_xf.set('applyNumberFormat', '1')
        # Also center-align (matches what the template does for the maturity col)
        align = etree.SubElement(new_xf, f'{{{NS_MAIN}}}alignment')
        align.set('horizontal', 'center')
        align.set('vertical', 'center')
        new_xf.set('applyAlignment', '1')
        cxfs.set('count', str(len(cxfs.findall('m:xf', NSMAP))))
        target_idx = len(cxfs.findall('m:xf', NSMAP)) - 1

    editor.files['xl/styles.xml'] = etree.tostring(
        root, xml_declaration=True, encoding='UTF-8', standalone=True)
    return str(target_idx)


def _build_changes_rows_xml(editor: _XlsxEditor, changes,
                              sheet_name='Changes this week', start_row=5):
    """Build XML for the 'Changes this week' sheet body (rows 5+).

    Reuses the template's existing cell styles so the output visually
    matches the published file:
      - Section label rows (Additions / Deletions / ...) use s="184" on
        the label cell and s="108" on the empty cells C..N, producing the
        gray-bar look.
      - Empty-state rows ("No changes this week") use s="185" / s="178".
      - Data rows use s="198" (and s="199" for the maturity cell) across
        the populated columns.

    Values are placed in columns 4 (D=ISIN), 6 (F=Issuer), 8 (H=Coupon),
    10 (J=Maturity), 12 (L=Previous rating), 14 (N=New rating) — the layout
    set up by the template's column headers on row 4.
    """
    # Style indexes borrowed from the template.
    # The template's data-row style 198 is bold, which reads poorly when
    # actual values are populated. We clone it (and 199 for dates) into
    # non-bold variants so the rows look like regular readable text.
    S_SECTION_LABEL = '184'   # gray-bar bold label — keep bold for headers
    S_SECTION_FILL  = '108'   # gray empty cell across the rest of the row
    S_DATA_LABEL    = '185'   # empty left-gutter cell on data rows
    S_DATA_FILL     = '178'   # empty-looking cell in the left-gutter columns
    S_DATA_VALUE    = _add_nonbold_variant(editor, 198)   # non-bold version
    S_DATA_DATE     = _add_nonbold_variant(editor, 199)   # non-bold date style

    parts = []
    row_num = start_row

    def blank_cells(row, cols, style):
        """Emit <c ... /> tags for a range of empty styled cells."""
        return ''.join(f'<c r="{c}{row}" s="{style}"/>' for c in cols)

    def add_section(label, items):
        nonlocal row_num
        # ----- Section label row: gray fill across B..N ----------------
        idx = editor.add_string(label)
        label_row_cells = (
            f'<c r="B{row_num}" s="{S_SECTION_LABEL}" t="s"><v>{idx}</v></c>'
            + blank_cells(row_num, 'CDEFGHIJKLMN', S_SECTION_FILL)
        )
        parts.append(
            f'<row r="{row_num}" ht="15" customHeight="1">'
            f'{label_row_cells}</row>'
        )
        row_num += 1

        if not items:
            empty_idx = editor.add_string('No changes this week')
            empty_row_cells = (
                f'<c r="B{row_num}" s="{S_DATA_LABEL}" t="s"><v>{empty_idx}</v></c>'
                + blank_cells(row_num, 'CDEFGHIJKLMN', S_DATA_FILL)
            )
            parts.append(
                f'<row r="{row_num}" ht="15" customHeight="1">'
                f'{empty_row_cells}</row>'
            )
            row_num += 2
            return

        for itm in items:
            # ISIN (D), Issuer (F), Coupon (H), Maturity (J), Prev (L), New (N)
            cells = [
                f'<c r="B{row_num}" s="{S_DATA_LABEL}"/>',
                f'<c r="C{row_num}" s="{S_DATA_FILL}"/>',
            ]
            isin_idx = editor.add_string(itm['isin'])
            cells.append(
                f'<c r="D{row_num}" s="{S_DATA_VALUE}" t="s"><v>{isin_idx}</v></c>')
            cells.append(f'<c r="E{row_num}" s="{S_DATA_VALUE}"/>')
            iss_idx = editor.add_string(itm['issuer'])
            cells.append(
                f'<c r="F{row_num}" s="{S_DATA_VALUE}" t="s"><v>{iss_idx}</v></c>')
            cells.append(f'<c r="G{row_num}" s="{S_DATA_VALUE}"/>')
            try:
                coupon_val = float(itm['coupon'].rstrip('%')) / 100 if itm.get('coupon') else None
            except (AttributeError, ValueError):
                coupon_val = None
            if coupon_val is not None:
                cells.append(
                    f'<c r="H{row_num}" s="{S_DATA_VALUE}"><v>{coupon_val}</v></c>')
            else:
                cells.append(f'<c r="H{row_num}" s="{S_DATA_VALUE}"/>')
            cells.append(f'<c r="I{row_num}" s="{S_DATA_VALUE}"/>')
            mat_dt = _parse_date(itm.get('maturity'))
            mat_text = (mat_dt.strftime('%m/%d/%Y') if mat_dt
                       else (str(itm.get('maturity') or '').strip() or 'Perpetual'))
            m_idx = editor.add_string(mat_text)
            cells.append(
               f'<c r="J{row_num}" s="{S_DATA_VALUE}" t="s"><v>{m_idx}</v></c>')
            cells.append(f'<c r="K{row_num}" s="{S_DATA_VALUE}"/>')
            prev_idx = editor.add_string(itm.get('view_prior') or '-')
            cells.append(
                f'<c r="L{row_num}" s="{S_DATA_VALUE}" t="s"><v>{prev_idx}</v></c>')
            cells.append(f'<c r="M{row_num}" s="{S_DATA_VALUE}"/>')
            new_idx = editor.add_string(itm.get('view_new') or '-')
            cells.append(
                f'<c r="N{row_num}" s="{S_DATA_VALUE}" t="s"><v>{new_idx}</v></c>')
            parts.append(
                f'<row r="{row_num}" ht="15" customHeight="1">'
                + ''.join(cells) + '</row>'
            )
            row_num += 1
        row_num += 1

    add_section('Additions',  changes['additions'])
    add_section('Deletions',  changes['deletions'])
    add_section('Upgrades',   changes['upgrades'])
    add_section('Downgrades', changes['downgrades'])
    return ''.join(parts)


# ════════════════════════════════════════════════════════════════════════════
# BUILDERS
# ════════════════════════════════════════════════════════════════════════════

def build_offshore_xlsx(data, output_path, template_path=DEFAULT_OFFSHORE_TEMPLATE):
    if not os.path.exists(template_path):
        raise FileNotFoundError(f'Offshore template not found: {template_path}')
    shutil.copyfile(template_path, output_path)
    os.chmod(output_path, 0o644)

    editor = _XlsxEditor(output_path)

    # Update cover-date cell
    sheet, row, col = OFFSHORE_COVER_DATE_CELL
    _set_cell_inline(editor, sheet, row, col, REPORT_DATE_FMT)

    # Unlock all sheets so compliance can edit
    _unlock_all_sheets(editor)

    # Replace the UBSLogo-font 'a'/'b' pair (renders as 'AB' on machines
    # without the proprietary UBSLogo font) with a plain 'UBS' text.
    #_replace_ubslogo_with_text(editor, replacement='UBS')

    # BondList
    def _issuer_maturity_sort_key(bond):
        issuer = data.issuer_display_name((bond.get('GK_Nummer') or '').strip(),
                                      fallback=bond.get('IssuerName', '')).lower()
        mat_raw = (bond.get('Maturity') or '').strip()
        mat_dt = _parse_date(mat_raw)
        # Perpetuals/missing dates sort last
        sort_dt = mat_dt if mat_dt else datetime.datetime(9999, 12, 31)
        isin = (bond.get('Isin') or '').strip()
        return (issuer, sort_dt, isin)
    
    rows = [b for b in data.em_bonds if is_offshore_eligible(b, data) and _has_issuer_name(b, data)]
    rows.sort(key=_issuer_maturity_sort_key)
    body_xml = _build_bondlist_rows_xml(editor, rows, data, 'offshore')
    _replace_sheet_data_rows(editor, 'BondList',
                              keep_first_n_rows=8, new_rows_xml=body_xml)
    # Always open on the Cover sheet, not wherever the template was last saved.
    _set_active_sheet(editor, 'Cover')
    editor.save()
    return len(rows)


def build_onshore_xlsx(data, output_path, template_path=DEFAULT_ONSHORE_TEMPLATE):
    if not os.path.exists(template_path):
        raise FileNotFoundError(f'Onshore template not found: {template_path}')
    shutil.copyfile(template_path, output_path)
    os.chmod(output_path, 0o644)

    editor = _XlsxEditor(output_path)

    sheet, row, col = ONSHORE_COVER_DATE_CELL
    _set_cell_inline(editor, sheet, row, col, REPORT_DATE_FMT)

    _unlock_all_sheets(editor)

    # Replace UBSLogo-font 'a'/'b' with plain 'UBS' text
    #_replace_ubslogo_with_text(editor, replacement='UBS')#

    # Onshore's Cover sheet has an extra 'ab' text in cell B1 (a second
    # fallback rendering of the UBS logo, this time via cell content not
    # a drawing). Overwrite it with 'UBS'.
    _set_cell_inline(editor, 'Cover', 1, 2, 'UBS')

    # BondList
    def _issuer_maturity_sort_key(bond):
        issuer = data.issuer_display_name((bond.get('GK_Nummer') or '').strip(),
                                      fallback=bond.get('IssuerName', '')).lower()
        mat_raw = (bond.get('Maturity') or '').strip()
        mat_dt = _parse_date(mat_raw)
        # Perpetuals/missing dates sort last
        sort_dt = mat_dt if mat_dt else datetime.datetime(9999, 12, 31)
        isin = (bond.get('Isin') or '').strip()
        return (issuer, sort_dt, isin)
    
    rows = [b for b in data.em_bonds if is_onshore_eligible(b, data) and _has_issuer_name(b, data)]
    rows.sort(key=_issuer_maturity_sort_key)
    body_xml = _build_bondlist_rows_xml(editor, rows, data, 'onshore')
    _replace_sheet_data_rows(editor, 'BondList',
                              keep_first_n_rows=8, new_rows_xml=body_xml)

    # ── Changes this week — filter the week-on-week diff to bonds that are
    #    US-onshore-eligible (this week for adds/upgrades/downgrades, last
    #    week for deletions). Previously this block was unreachable because
    #    an early editor.save()/return sat above it, so the sheet kept its
    #    template "No changes this week" placeholder. ──────────────────────
    def _onshore_change(row):
        isin = (row.get('isin') or '').strip()
        if not isin:
            return False
        # Current week (additions, upgrades, downgrades)
        bond = data.bond_by_isin.get(isin)
        if bond is not None and is_onshore_eligible(bond, data):
            return True
        # Previous week (deletions) — check against last week's data
        prev_bond = getattr(data, 'prev_bonds', {}).get(isin)
        if not prev_bond:
            return False
        prev_upd = getattr(data, 'prev_bond_updates', {}).get(isin, {})
        return is_onshore_eligible(prev_bond, data, upd=prev_upd)

    if 'Changes this week' in editor._sheet_map:
        changes = data.recommendation_changes()
        changes = {k: [r for r in v if _onshore_change(r)] for k, v in changes.items()}
        chg_xml = _build_changes_rows_xml(editor, changes)
        _replace_sheet_data_rows(editor, 'Changes this week',
                                  keep_first_n_rows=4, new_rows_xml=chg_xml)

    # NOTE: the original design also wiped 'Issuer rating history' and the
    # 'Disclosures' sheets to a single placeholder note. That is intentionally
    # left DISABLED: in the current onshore template these sheets ship with
    # real content (the rating-history sheet has 360+ populated rows), and
    # wiping them would destroy it. Re-enable explicitly only if the weekly
    # process is supposed to clear them for downstream teams to refill.
    # if 'Issuer rating history' in editor._sheet_map:
    #     _wipe_sheet_to_note(editor, 'Issuer rating history',
    #                          'Populated weekly by the rating-history team.')
    # for name in ('Disclosures', 'Disclosures (2)'):
    #     if name in editor._sheet_map:
    #         _wipe_sheet_to_note(editor, name,
    #                              'Populated manually each week before distribution.')

    # Always open on the Cover sheet, not wherever the template was last saved.
    _set_active_sheet(editor, 'Cover')
    editor.save()
    return len(rows)


def build_excels(data, offshore_path=None, onshore_path=None,
                  offshore_template=DEFAULT_OFFSHORE_TEMPLATE,
                  onshore_template=DEFAULT_ONSHORE_TEMPLATE):
    out = {}
    if offshore_path:
        n = build_offshore_xlsx(data, offshore_path, offshore_template)
        out['offshore'] = n
        print(f'[xlsx] wrote {offshore_path}  ({n:,} bonds)')
    if onshore_path:
        n = build_onshore_xlsx(data, onshore_path, onshore_template)
        out['onshore'] = n
        print(f'[xlsx] wrote {onshore_path}  ({n:,} bonds)')
    return out
