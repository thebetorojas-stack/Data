"""haver_smoke_test.py — confirms Python ↔ Haver wiring is working.

Runs four checks:
  1. Can we import the Haver module at all?
  2. Which databases does this seat have access to?
  3. Can we fetch a known sentinel series (US GDP from USECON)?
  4. Are Chile and Argentina databases available, and what codes do they use?

Usage:  python haver_smoke_test.py

If you get ModuleNotFoundError, edit DLX_PATH below to point at your DLX folder
(default is C:\\DLX\\Data), or set the PYTHONPATH environment variable:
    setx PYTHONPATH "C:\\DLX\\Data"
then close and reopen the command prompt.

Works on Python 3.8+.
"""

from __future__ import annotations

import os
import sys
from datetime import date, timedelta

# Add common DLX install paths to sys.path before trying to import Haver.
DLX_CANDIDATES = [
    r"C:\DLX\Data",
    r"D:\DLX\Data",
    r"C:\Program Files\DLX\Data",
    r"C:\Program Files (x86)\DLX\Data",
]
for p in DLX_CANDIDATES:
    if os.path.isdir(p) and p not in sys.path:
        sys.path.insert(0, p)


# ─────────────────────────────────────────────────────────────────────────── #
def main() -> int:
    print("─" * 70)
    print("Haver smoke test")
    print("─" * 70)

    # ── Check 1: import ──
    try:
        import Haver  # type: ignore
    except ImportError as e:
        print(f"\n[FAIL] step 1 — cannot import Haver: {e}")
        print("\nFix:")
        print("  Verify DLX is installed (look for C:\\DLX\\Data\\Haver.py).")
        print("  Add the folder to PYTHONPATH:  setx PYTHONPATH \"C:\\DLX\\Data\"")
        print("  Close the command prompt and open a new one, then retry.")
        return 1
    print(f"[ OK ] step 1 — imported Haver from {getattr(Haver, '__file__', '?')}")

    # ── Check 2: list databases ──
    print()
    try:
        # Some Haver versions expose path() to set the data folder; call it
        # defensively in case it's required.
        for p in DLX_CANDIDATES:
            if os.path.isdir(p):
                try:
                    Haver.path(p)
                except Exception:
                    pass
                break
        dbs = list(Haver.databases())
    except Exception as e:
        print(f"[WARN] step 2 — could not list databases: {e}")
        dbs = []
    if dbs:
        print(f"[ OK ] step 2 — {len(dbs)} databases available")
        # show first ~30 to keep output readable
        preview = ", ".join(sorted(dbs)[:30])
        more = "" if len(dbs) <= 30 else f"  (+ {len(dbs) - 30} more)"
        print(f"        databases: {preview}{more}")
    else:
        print("[WARN] step 2 — database list empty; some Haver versions don't expose this")

    # ── Check 3: sentinel pull (US GDP from USECON) ──
    print()
    try:
        end = date.today()
        start = end - timedelta(days=400)
        df = Haver.data(["gdp"], database="USECON",
                        startdate=start.isoformat(),
                        enddate=end.isoformat())
        if df is None or len(df) == 0:
            print("[WARN] step 3 — sentinel pull returned empty (do you have USECON?)")
        else:
            last_idx = df.index[-1]
            last_val = df.iloc[-1, 0]
            print(f"[ OK ] step 3 — fetched {len(df)} rows of USECON.gdp")
            print(f"        last obs: {last_idx} = {last_val}")
    except Exception as e:
        print(f"[FAIL] step 3 — sentinel pull failed: {e}")
        print("        This usually means USECON isn't in your subscription.")
        print("        Try the next step to find what databases you DO have.")

    # ── Check 4: Chile / Argentina databases ──
    print()
    candidates_cl = ["CHILE", "CHL", "CHILEDB", "LATIN", "EMERGE"]
    candidates_ar = ["ARGENT", "ARG", "ARGENTNA", "ARGENTINA", "LATIN", "EMERGE"]

    found_cl = [c for c in candidates_cl if dbs and c in dbs]
    found_ar = [c for c in candidates_ar if dbs and c in dbs]

    if dbs:
        if found_cl:
            print(f"[ OK ] Chile candidates present: {', '.join(found_cl)}")
        else:
            print("[WARN] None of the standard Chile database codes found.")
            print(f"        Expected one of: {candidates_cl}")
        if found_ar:
            print(f"[ OK ] Argentina candidates present: {', '.join(found_ar)}")
        else:
            print("[WARN] None of the standard Argentina database codes found.")
            print(f"        Expected one of: {candidates_ar}")

    # Try fetching a sentinel series from each found candidate to confirm
    print()
    sentinel_codes = ["c917cpi", "c213cpi", "cpi"]   # try a few common codes
    for db in found_cl + found_ar:
        for code in sentinel_codes:
            try:
                df = Haver.data([code], database=db,
                                startdate="2024-01-01",
                                enddate=date.today().isoformat())
                if df is not None and len(df) > 0:
                    last = df.iloc[-1, 0]
                    print(f"[ OK ] {db}.{code} → {len(df)} obs, last = {last}")
                    break
            except Exception:
                continue
        else:
            print(f"[INFO] {db} accessible but couldn't find any of {sentinel_codes}")
            print(f"        (try Haver DLX UI to browse {db} for the right code)")

    print()
    print("─" * 70)
    print("Done. If steps 1–3 passed, you're hooked up. Tell me which databases")
    print("you have for Chile and Argentina, and I'll wire them into the model.")
    print("─" * 70)
    return 0


if __name__ == "__main__":
    sys.exit(main())
