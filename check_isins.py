#!/usr/bin/env python3
"""
check_isins.py — "why are these ISINs missing?" -> one Excel with the answers.

HOW TO USE (3 steps):
  1. Open  isins_to_check.txt , paste the ISINs people asked about (one per
     line), save it.
  2. Run:  python check_isins.py
  3. Open  outputs/why_missing.xlsx  -> one row per ISIN with the reason.

It uses the SAME eligibility logic as the real weekly build, so its answers
always match the published PDF / Offshore / Onshore outputs.
"""
import os
import sys
import datetime

import gem_report_builder_v3 as G
from gem_excel_builder import is_onshore_eligible, is_offshore_eligible, _has_issuer_name

HERE = os.path.dirname(os.path.abspath(__file__))
CURR = os.path.join(HERE, 'data', 'current')
PREV = os.path.join(HERE, 'data', 'previous')
ISIN_FILE  = os.path.join(HERE, 'isins_to_check.txt')
OUT_DIR    = os.path.join(HERE, 'outputs')
OUT_XLSX   = os.path.join(OUT_DIR, 'why_missing.xlsx')


# Plain-English explanation for each master-list reason code.
EXPLAIN = {
    'legal_excluded':       'Legal asked us to pull it (on the legal exclusions list).',
    'not_gem':              "Not tagged into the GEM universe (TopListCategory is not 'GEM' in the bond-update file). Most common cause.",
    'near_maturity':        'Matures within 180 days, so it drops off (except defaulted Venezuelan sovereigns).',
    'subordinated_no_flag': 'Subordinated bond with no WMRFlag and no usable rating, so it is flagged, not included.',
    'eligible':             'Eligible.',
    'subordinated_ig':      'Eligible (subordinated, investment grade).',
    'subordinated_hy':      'Eligible (subordinated, high yield).',
    'subordinated_unrated': 'Eligible (subordinated, unrated - included as HY).',
}


def load_isins():
    if len(sys.argv) > 1:                     # ISINs passed on the command line win
        return [a.strip().upper() for a in sys.argv[1:] if a.strip()]
    if not os.path.exists(ISIN_FILE):
        sys.exit(f"Can't find {ISIN_FILE}. Create it and paste ISINs, one per line.")
    out = []
    for line in open(ISIN_FILE, encoding='utf-8'):
        line = line.strip()
        if line and not line.startswith('#'):
            out.append(line.upper())
    if not out:
        sys.exit(f"No ISINs found in {ISIN_FILE}. Paste some in (one per line) and re-run.")
    return out


def build_data():
    return G.GEMData({
        'bond_data':      os.path.join(CURR, 'CurrentPublishableBondData.txt'),
        'issuer_data':    os.path.join(CURR, 'CurrentPublishableIssuerData.txt'),
        'bond_update':    os.path.join(CURR, 'PublishableBondDataUpdate.txt'),
        'issuer_update':  os.path.join(CURR, 'PublishableIssuerDataUpdate.txt'),
        'color_flags':    os.path.join(CURR, 'PublishableColorFlags.txt'),
        'issuer_texts':   os.path.join(CURR, 'IssuerTexts.txt'),
        'issuer_ratings': os.path.join(CURR, 'IssuerRatings.txt'),
        'prev_bond_data':   os.path.join(PREV, 'CurrentPublishableBondData.txt'),
        'prev_bond_update': os.path.join(PREV, 'PublishableBondDataUpdate.txt'),
        'priips_ref':       None,
        'legal_exclusions': _find_legal_exclusions(),
    })


def _find_legal_exclusions():
    for folder in (os.path.join(HERE, 'data'), CURR, HERE):
        if os.path.isdir(folder):
            for f in os.listdir(folder):
                lf = f.lower()
                if lf.endswith(('.csv', '.txt')) and ('legal' in lf or 'exclusion' in lf) \
                        and not f.startswith('~$'):
                    return os.path.join(folder, f)
    return None


def assess(isin, data):
    """Return a dict of answers for one ISIN."""
    r = {'ISIN': isin, 'In feed?': 'NO', 'On the list?': 'NO',
         'Reason': '', 'Explanation': '', 'Offshore Excel?': '', 'Onshore Excel?': ''}

    bond = data.bond_by_isin.get(isin)
    if bond is None:
        r['Reason'] = 'not_in_feed'
        r['Explanation'] = ("This ISIN is not in this week's bond feed at all. "
                            "It never reached us - check the ISIN is correct, or the "
                            "desk needs to add it upstream. Nothing this tool can include.")
        return r

    r['In feed?'] = 'YES'
    upd = data.bond_updates.get(isin, {})
    gk  = (bond.get('GK_Nummer') or '').strip()
    r['Issuer'] = data.issuer_display_name(gk, fallback=bond.get('IssuerName', ''))

    decision = data._classify_for_list(bond, upd)
    r['Reason'] = decision['reason']
    r['Explanation'] = EXPLAIN.get(decision['reason'], decision['reason'])

    if decision['eligible']:
        r['On the list?'] = 'YES'
        r['Offshore Excel?'] = 'YES' if is_offshore_eligible(bond, data) else 'NO'
        on = is_onshore_eligible(bond, data) and _has_issuer_name(bond, data)
        r['Onshore Excel?'] = 'YES' if on else 'NO'
        if r['Offshore Excel?'] == 'YES' and r['Onshore Excel?'] == 'YES':
            r['Explanation'] = 'It IS on the list (offshore and onshore). If someone says it is missing, check the sheet/tab they are looking at.'
        elif r['Onshore Excel?'] == 'NO':
            r['Explanation'] = ('On the master list and Offshore, but NOT on the Onshore Excel '
                                '(usually: not USD, or a Reg-S / 144A corporate, or an excluded-country sovereign).')
    return r


def write_xlsx(rows):
    os.makedirs(OUT_DIR, exist_ok=True)
    cols = ['ISIN', 'Issuer', 'In feed?', 'On the list?',
            'Offshore Excel?', 'Onshore Excel?', 'Reason', 'Explanation']
    try:
        import openpyxl
        from openpyxl.styles import Font, PatternFill, Alignment
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = 'Why missing'
        ws.append(cols)
        for c in ws[1]:
            c.font = Font(bold=True, color='FFFFFF')
            c.fill = PatternFill('solid', fgColor='305496')
        for row in rows:
            ws.append([row.get(c, '') for c in cols])
            status = row.get('On the list?')
            fill = '006100' if status == 'YES' else '9C0006'
            fillbg = 'C6EFCE' if status == 'YES' else 'FFC7CE'
            cell = ws.cell(row=ws.max_row, column=4)
            cell.fill = PatternFill('solid', fgColor=fillbg)
            cell.font = Font(color=fill, bold=True)
        widths = [16, 30, 9, 13, 16, 16, 22, 70]
        for i, w in enumerate(widths, 1):
            ws.column_dimensions[openpyxl.utils.get_column_letter(i)].width = w
        for r_ in ws.iter_rows(min_row=2):
            r_[-1].alignment = Alignment(wrap_text=True, vertical='top')
        ws.freeze_panes = 'A2'
        wb.save(OUT_XLSX)
        return OUT_XLSX
    except ImportError:
        import csv
        path = OUT_XLSX.replace('.xlsx', '.csv')
        with open(path, 'w', newline='', encoding='utf-8-sig') as f:
            w = csv.DictWriter(f, fieldnames=cols)
            w.writeheader()
            for row in rows:
                w.writerow({c: row.get(c, '') for c in cols})
        return path


def main():
    isins = load_isins()
    print(f'Checking {len(isins)} ISIN(s)...')
    data = build_data()
    rows = [assess(i, data) for i in isins]
    path = write_xlsx(rows)
    print('\nDone. Answers written to:')
    print('   ' + path)
    print('\nQuick summary:')
    for r in rows:
        print(f"   {r['ISIN']:<14} {r['On the list?']:<4} {r['Reason']}")


if __name__ == '__main__':
    main()
