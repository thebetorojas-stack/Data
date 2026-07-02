#!/usr/bin/env python3
"""
why_missing.py — answer "why didn't this ISIN show up?" in one shot.

It builds the EXACT same GEMData the weekly run builds, then walks a bond
through every gate that can drop it — for the master list, the Offshore
Excel, and the Onshore Excel — and tells you which gate killed it (or
confirms it's actually there).

Because it calls the real classifier functions (_classify_for_list,
is_offshore_eligible, is_onshore_eligible), it can NEVER disagree with the
published outputs.

USAGE
  Command line (fastest, no console):
      python why_missing.py XS1234567890
      python why_missing.py XS1234567890 US40012345AA XS999...   # several at once

  Python console (if you want to poke around after):
      from why_missing import build, why
      data = build()          # loads this week's data once
      why("XS1234567890", data)
      why("XS1234567890", data)   # reuse `data`, it's the slow part
"""
import os
import sys
import datetime

import gem_report_builder_v3 as G
from gem_excel_builder import (
    is_onshore_eligible, is_offshore_eligible, _has_issuer_name,
    _is_reg_s, _is_quasi_sovereign,
    ONSHORE_EXCLUDE_OVERRIDES, ONSHORE_INCLUDE_OVERRIDES,
    ONSHORE_SOV_EXCLUDED_COUNTRIES,
)

# Same layout run_weekly.py uses.
HERE = os.path.dirname(os.path.abspath(__file__))
CURR = os.path.join(HERE, 'data', 'current')
PREV = os.path.join(HERE, 'data', 'previous')


def build():
    """Load this week's data exactly like run_weekly.py does."""
    paths = {
        'bond_data':      os.path.join(CURR, 'CurrentPublishableBondData.txt'),
        'issuer_data':    os.path.join(CURR, 'CurrentPublishableIssuerData.txt'),
        'bond_update':    os.path.join(CURR, 'PublishableBondDataUpdate.txt'),
        'issuer_update':  os.path.join(CURR, 'PublishableIssuerDataUpdate.txt'),
        'color_flags':    os.path.join(CURR, 'PublishableColorFlags.txt'),
        'issuer_texts':   os.path.join(CURR, 'IssuerTexts.txt'),
        'issuer_ratings': os.path.join(CURR, 'IssuerRatings.txt'),
        'prev_bond_data':   os.path.join(PREV, 'CurrentPublishableBondData.txt'),
        'prev_bond_update': os.path.join(PREV, 'PublishableBondDataUpdate.txt'),
        'priips_ref':       None,   # Restrictions column only; not an eligibility gate
        'legal_exclusions': _find_legal_exclusions(),
    }
    return G.GEMData(paths)


def _find_legal_exclusions():
    """Mirror run_weekly's search so legal pulls are reflected here too."""
    for folder in (os.path.join(HERE, 'data'), CURR, HERE):
        if not os.path.isdir(folder):
            continue
        for f in os.listdir(folder):
            lf = f.lower()
            if (lf.endswith(('.csv', '.txt'))
                    and ('legal' in lf or 'exclusion' in lf)
                    and not f.startswith('~$')):
                return os.path.join(folder, f)
    return None


def why(isin, data):
    """Print the full decision trail for one ISIN. Returns nothing; just prints."""
    isin = (isin or '').strip().upper()
    print('=' * 68)
    print(f'ISIN: {isin}')
    print('=' * 68)

    # ---- Gate 0: is it even in the bank's bond feed this week? -------------
    bond = data.bond_by_isin.get(isin)
    if bond is None:
        # try a case-insensitive / whitespace match to catch feed quirks
        for k, b in data.bond_by_isin.items():
            if k.strip().upper() == isin:
                bond, isin = b, k
                break
    if bond is None:
        print('NOT IN THE BOND FEED at all this week.')
        print('  -> It never reached us. Check CurrentPublishableBondData.txt.')
        print('     Reasons: not in the WM universe, ISIN typo, or the desk')
        print('     needs to add it upstream. Nothing this tool can include.')
        return

    upd = data.bond_updates.get(isin, {})
    gk  = (bond.get('GK_Nummer') or '').strip()
    print(f"In feed. Issuer: {data.issuer_display_name(gk, fallback=bond.get('IssuerName',''))}")
    print(f"  GK={gk}  CCY={bond.get('CCY','')}  Maturity={bond.get('Maturity','')}")
    print(f"  TopListCategory={upd.get('TopListCategory','(none)')!r}  "
          f"WMRFlag={(bond.get('WMRFlag') or upd.get('WMRFlag') or '')!r}  "
          f"subordinated={G.is_subordinated_bond(bond)}")
    has_update_row = bool(upd)
    if not has_update_row:
        print("  !! No PublishableBondDataUpdate.txt row -> TopListCategory is")
        print("     blank, so it fails the GEM check below. This is the single")
        print("     most common cause. Check the update file.")
    print('-' * 68)

    # ---- MASTER LIST (drives the PDF and both Excels) ---------------------
    decision = data._classify_for_list(bond, upd)
    reason_help = {
        'legal_excluded':       'Legal asked us to pull it (legal exclusions file / issuer name).',
        'not_gem':              "TopListCategory != 'GEM' in the bond-update file. Not tagged into the universe.",
        'near_maturity':        'Matures within 180 days (and not a defaulted VE sovereign).',
        'subordinated_no_flag': 'Subordinated with no WMRFlag and no usable rating -> flagged, not included.',
    }
    print(f"MASTER LIST  -> eligible={decision['eligible']}  reason={decision['reason']}")
    if not decision['eligible']:
        print('   X ' + reason_help.get(decision['reason'], decision['reason']))
        print('   -> If it fails here, it is absent from the PDF AND both Excels.')
        return
    in_master = isin in {(b.get('Isin') or '').strip() for b in data.em_bonds}
    print(f'   OK on master list (in data.em_bonds = {in_master}).')
    print('-' * 68)

    # ---- OFFSHORE EXCEL ---------------------------------------------------
    off = is_offshore_eligible(bond, data)
    print(f'OFFSHORE Excel -> {off}')
    if not off:
        print('   X Failed offshore gate (GEM tag or 180-day maturity).')

    # ---- ONSHORE EXCEL (the pickiest one) --------------------------------
    on = is_onshore_eligible(bond, data)
    print(f'ONSHORE Excel  -> {on}')
    if not on:
        print('   X Failed onshore gate. Checking which rule:')
        ccy = (bond.get('CCY') or '').strip().upper()
        if ccy != 'USD':
            print(f'      - not USD (CCY={ccy!r}). Onshore is USD-only.')
        elif (upd.get('TopListCategory') or '').strip() != 'GEM':
            print('      - not tagged GEM.')
        elif isin in ONSHORE_EXCLUDE_OVERRIDES:
            print('      - on ONSHORE_EXCLUDE_OVERRIDES (hard-coded force-exclude).')
        else:
            itype = data.issuer_type(gk)
            cc    = data.issuer_country_code(gk)
            name  = data.issuer_display_name(gk, fallback=bond.get('IssuerName', ''))
            if itype == 'SOV' and cc in ONSHORE_SOV_EXCLUDED_COUNTRIES:
                print(f'      - sovereign of excluded country {cc}.')
            elif itype in ('SOV', 'SUPRA') or _is_quasi_sovereign(gk, name):
                print(f'      - (itype={itype}) should normally be eligible; check maturity/overrides.')
            elif _is_reg_s(bond):
                print(f'      - Reg S corporate/financial (CINS prefix or Market="International").')
                print(f'        144A/Reg-S paper is kept OFF the onshore list by design.')
            else:
                print(f'      - itype={itype}; failed for another reason, inspect the bond dict.')
    elif on and not _has_issuer_name(bond, data):
        print('   ! onshore-eligible but has NO issuer name -> dropped from the sheet rows.')

    print('=' * 68)


def main(argv):
    if not argv:
        print(__doc__)
        return
    data = build()
    for isin in argv:
        why(isin, data)
        print()


if __name__ == '__main__':
    main(sys.argv[1:])
