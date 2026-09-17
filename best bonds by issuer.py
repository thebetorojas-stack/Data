# -*- coding: utf-8 -*-
"""
best_bonds_by_issuer.py
=======================
Same analysis as best_bonds_by_country.py, for ANY issuer.

  1. Type the issuer below (part of the name is enough, case doesn't matter).
  2. F5.
  -> outputs/Best_Bonds_<Issuer>_<date>.xlsx

CLI:  python best_bonds_by_issuer.py --issuer "Petrobras"
      python best_bonds_by_issuer.py --list "pemex"     (show matching issuer names)

Keep this file in the same folder as best_bonds_by_country.py — the scoring
engine and all settings (peers, scenarios, weights) live there.
"""

# =============================================================================
ISSUER = "Petroleos Mexicanos"      # <-- CHANGE THIS
# =============================================================================

EXACT_MATCH = False                 # True = full name must match exactly

# Short names people actually use -> wording in the feed's issuer name
ISSUER_ALIASES = {
    "pemex": "Petroleos Mexicanos",
    "cfe": "Comision Federal de Electricidad",
    "petrobras": "Petrobras",
    "ecopetrol": "Ecopetrol",
    "codelco": "Codelco",
    "ypf": "YPF",
    "argentina": "Republic of Argentina",
    "mexico": "United Mexican States",
    "brazil": "Federative Republic of Brazil",
    "chile": "Republic of Chile",
}
DATA_DIR = None                     # None = use DATA_DIR from best_bonds_by_country.CONFIG
CURRENCY = None                     # None = same as the country script ("USD")

import argparse
import sys

import best_bonds_by_country as eng


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--issuer")
    ap.add_argument("--list", metavar="TEXT")
    ap.add_argument("--data")
    a = ap.parse_args()

    cfg = dict(eng.CONFIG)
    cfg["UNIVERSE"] = "all"           # an issuer can be a corporate or quasi
    if DATA_DIR:
        cfg["DATA_DIR"] = DATA_DIR
    if a.data:
        cfg["DATA_DIR"] = a.data
    if CURRENCY:
        cfg["CURRENCY"] = CURRENCY
    issuer = (a.issuer or ISSUER).strip()
    issuer = ISSUER_ALIASES.get(issuer.lower(), issuer)

    if a.list is not None:
        raw, _, _ = eng.load_universe(cfg)
        names = sorted({str(x) for x in raw["issuer"].dropna() if a.list.lower() in str(x).lower()})
        print("\n".join(names) or "no match")
        return

    raw, _, _ = eng.load_universe(cfg, verbose=False)
    all_names = sorted({str(x).strip() for x in raw["issuer"].dropna()})
    if EXACT_MATCH:
        matches = [n for n in all_names if n.lower() == issuer.lower()]
    else:
        matches = [n for n in all_names if issuer.lower() in n.lower()]
    if not matches:
        import difflib
        close = [n for n in all_names if any(w in n.lower() for w in issuer.lower().split() if len(w) > 3)]
        close += [n for n in difflib.get_close_matches(issuer, all_names, n=8, cutoff=0.4) if n not in close]
        close += [n for n in all_names if n not in close and
                  any(difflib.SequenceMatcher(None, issuer.lower(), w).ratio() > 0.8 for w in n.lower().split())]
        print(f"No issuer matches '{issuer}'." + (" Did you mean:\n  " + "\n  ".join(close[:15]) if close else ""))
        sys.exit(1)
    print(f"Issuer names matched: {', '.join(matches)}")

    wanted = {m.lower() for m in matches}
    out, results, picks = eng.run(cfg, lambda s: s.issuer.str.lower().isin(wanted), issuer)
    if len(matches) > 1:
        print("  Note: several issuer names matched — each gets its own curve in the file. "
              "Set EXACT_MATCH = True to narrow.")


if __name__ == "__main__":
    main()
