"""Weekly runner for the EM Bond List pipeline.

Open this file in Spyder and press F5. Any errors will show clearly in
the console. No subprocess wrapper hiding the real problem.
"""

import os
import sys
import glob

# Make sure this script runs from its own folder, so relative paths work
HERE = os.path.dirname(os.path.abspath(__file__))
os.chdir(HERE)
if HERE not in sys.path:
    sys.path.insert(0, HERE)

DATA = 'data'
CURR = os.path.join(DATA, 'current')
PREV = os.path.join(DATA, 'previous')
OUT  = 'outputs'

os.makedirs(OUT, exist_ok=True)


def find_priips_reference():
    """Locate the authoritative PRIIPS / MiFID classification extract.

    The India team keeps this file up to date in the `data/` folder (alongside
    the `current/` and `previous/` sub-folders, NOT inside them). We DON'T
    require it to be renamed — we pick up the newest file whose name contains
    'priips'. Search order: data/ first (the agreed home), then data/current
    and the folder root as fall-backs. Returns the path, or None if not found.

    Why this file is required and not computed: the MiFID-complexity (code 1)
    and PRIIPS-relevant/KID (code 2) flags are per-instrument compliance
    determinations. They are NOT derivable from the bond feed — verified on the
    extract itself, plain vanilla bonds (incl. identical US Treasuries) split
    ~55/45 complex vs non-complex. So the data must come from the source; only
    its *delivery* and *application* are automated.
    """
    patterns = []
    for folder in (DATA, CURR, HERE):           # data/ is the agreed location
        for ext in ('xls', 'xlsx', 'csv'):
            patterns.append(os.path.join(folder, '*[Pp][Rr][Ii][Ii][Pp][Ss]*.' + ext))
    candidates = []
    for pat in patterns:
        candidates.extend(glob.glob(pat))
    # De-dupe, drop Excel lock files (~$...), pick most recently modified
    candidates = [c for c in set(candidates)
                  if not os.path.basename(c).startswith('~$')]
    if not candidates:
        return None
    return max(candidates, key=os.path.getmtime)


priips_ref = find_priips_reference()

# Simulate the command-line arguments that gem_report_builder_v3.main() expects
sys.argv = [
    'gem_report_builder_v3.py',
    '--bond-data',        os.path.join(CURR, 'CurrentPublishableBondData.txt'),
    '--issuer-data',      os.path.join(CURR, 'CurrentPublishableIssuerData.txt'),
    '--bond-update',      os.path.join(CURR, 'PublishableBondDataUpdate.txt'),
    '--issuer-update',    os.path.join(CURR, 'PublishableIssuerDataUpdate.txt'),
    '--color-flags',      os.path.join(CURR, 'PublishableColorFlags.txt'),
    '--issuer-texts',     os.path.join(CURR, 'IssuerTexts.txt'),
    '--issuer-ratings',   os.path.join(CURR, 'IssuerRatings.txt'),
    '--prev-bond-data',   os.path.join(PREV, 'CurrentPublishableBondData.txt'),
    '--prev-bond-update', os.path.join(PREV, 'PublishableBondDataUpdate.txt'),
    '--output',           os.path.join(OUT,  'GEM_List.pdf'),
    '--xlsx-offshore',    os.path.join(OUT,  'GEM_List_Offshore.xlsx'),
    '--xlsx-onshore',     os.path.join(OUT,  'GEM_List_Onshore.xlsx'),
    '--ladder-output',    os.path.join(OUT,  'LatAm_Bond_Ladder.pdf'),
]

if priips_ref:
    sys.argv += ['--priips-ref', priips_ref]
    print('Using PRIIPS reference:', priips_ref)
else:
    print('!! WARNING: no PRIIPS reference file found in the', DATA, 'folder.\n'
          '   The India team must keep the latest extract (any name containing '
          '"PRIIPS",\n   .xls/.xlsx/.csv) in the data/ folder. Until then, '
          'restrictions fall back to\n   the OLD heuristic (less accurate).')

print('\nStarting pipeline — should take ~30 seconds.\n')

import gem_report_builder_v3
data = gem_report_builder_v3.main()

# Surface how the restrictions column was sourced this run.
n_ref = len(getattr(data, 'priips_ref', {}) or {})
n_unmatched = len({row[0] for row in getattr(data, 'priips_unmatched', [])})
if n_ref:
    print(f'\nRestrictions: looked up from PRIIPS reference '
          f'({n_ref:,} instruments). '
          f'{n_unmatched} bond(s) on the list were not in the extract '
          f'(shown as "n/a"; see outputs/priips_unmatched_report.csv).')

print('\n✓ Done. Files in', OUT + ':')
for f in sorted(os.listdir(OUT)):
    print('   ', f)

import gem_excel_builder
gem_excel_builder.build_excels(data, offshore_path=os.path.join(OUT, 'GEM_List_Offshore.xlsx'), onshore_path=os.path.join(OUT, 'GEM_List_Onshore.xlsx')),
import gem_ladder_builder
gem_ladder_builder.build_ladder_pdf(data, os.path.join(OUT, 'LatAm_Bond_Ladder.pdf'))

# Clean, shareable restrictions workbook (instructions for India + the current
# Offshore restrictions). Regenerated every run from the freshly-built Offshore
# list, so it always reflects this week's bonds — additions and removals
# included — with zero manual upkeep.
import restrictions_clean_builder
_clean_n = restrictions_clean_builder.build_restrictions_clean_xlsx(
    data,
    offshore_path=os.path.join(OUT, 'GEM_List_Offshore.xlsx'),
    output_path=os.path.join(OUT, 'EMBL_Restrictions_CLEAN.xlsx'))
print(f'[clean] wrote EMBL_Restrictions_CLEAN.xlsx ({_clean_n:,} offshore bonds)')
