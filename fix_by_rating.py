import sys
P = sys.argv[1] if len(sys.argv) > 1 else "embi_builder.py"
src=open(P,encoding="utf-8").read()
def rep(old, new, n=1):
    """Replace an exact block, refusing loudly rather than half-applying."""
    global src
    if src.count(old) != n:
        if new.strip() and new[:60] in src:
            sys.exit("Already applied - nothing to do.")
        sys.exit("NO MATCH (found %d, expected %d) for:\n%s\n\n"
                 "Your file is not the version this patch expects. Stop and re-sync it."
                 % (src.count(old), n, old[:200]))
    src = src.replace(old, new, n)

# ---- 1. metrics for the published YTD attribution ----------------------
rep('''    "yield":  "Yld to Maturity",
    "tret":   "Cum Tot Ret Idx",
}''',
'''    "yield":  "Yld to Maturity",
    "tret":   "Cum Tot Ret Idx",
    # Published YTD attribution, carried straight off the snapshot rather than
    # derived from a return index. JPM ships these per country AND per rating
    # bucket; earlier builds discarded them, which is why By_Rating had no YTD.
    "tr_ytd":     "TR YTD (%)",
    "spr_ytd":    "Spread Ret YTD (%)",
    "ust_ytd":    "UST Ret YTD (%)",
}''')

# ---- 2. rating bucket name aliases -------------------------------------
rep('''# S&P / Moody-style rating to numeric score (higher = better quality).''',
'''# The JPM "regional" export names its credit buckets differently from the
# returns file. RATING_ORDER (and therefore every rating block in the workbook)
# is keyed on the returns-file spelling, so without this map the snapshot's
# buckets simply never match and By_Rating renders empty.
RATING_BUCKET_ALIASES: Dict[str, str] = {
    "investment grade":         "Credit IG only",
    "non investment grade":     "Credit Non-IG",
    "non-investment grade":     "Credit Non-IG",
    "aa":                       "Credit AA only",
    "a":                        "Credit A only",
    "bbb":                      "Credit BBB only",
    "bb":                       "Credit BB only",
    "b":                        "Credit B only",
    "c":                        "Credit C only",
    "nr":                       "Credit NR",
    "residual":                 "Credit Residual only",
}

# S&P / Moody-style rating to numeric score (higher = better quality).''')

# ---- 3. load_snapshot: strip header whitespace + map buckets -----------
rep('''    with _open_csv(path) as f:
        reader = csv.DictReader(f)
        rows = list(reader)
    snap_date: Optional[datetime] = None''',
'''    with _open_csv(path) as f:
        reader = csv.DictReader(f)
        # JPM ships several headers with a trailing space
        # ("UST Return YTD Change (%) "). Strip every key once here so callers
        # can look columns up by their clean name.
        rows = [{(k or "").strip(): v for k, v in r.items()} for r in reader]
    snap_date: Optional[datetime] = None''')

rep('''        elif instrument in ("EMBI Global", "EMBI Global Diversified"):''',
'''        elif instrument.strip().lower() in RATING_BUCKET_ALIASES:
            # 'Investment Grade', 'BB', 'NR' ... -> the returns-file spelling
            # the rest of the workbook keys on.
            key = RATING_BUCKET_ALIASES[instrument.strip().lower()]
        elif instrument in ("EMBI Global", "EMBI Global Diversified"):''')

# ---- 4. ingest the YTD columns -----------------------------------------
rep('''    "STW (Trsy)":                "STW (Trsy)",
}

# Every spelling under which a per-country spread can arrive, best first.''',
'''    "STW (Trsy)":                "STW (Trsy)",
    # Published YTD attribution (percent units, e.g. -0.52 = -0.52%).
    "YTD Change (%)":                "TR YTD (%)",
    "Spread Return YTD Change (%)":  "Spread Ret YTD (%)",
    "UST Return YTD Change (%)":     "UST Ret YTD (%)",
}

# Every spelling under which a per-country spread can arrive, best first.''')

# ---- 5. keep rating buckets out of the country rollups -----------------
rep('''        region = COUNTRY_REGION.get(ent)
        if region is None:
            continue
        if EXCLUDE_DEFAULTED_FROM_AGGREGATE and ent in DEFAULTED_COUNTRIES:''',
'''        region = COUNTRY_REGION.get(ent)
        if region is None:
            continue  # rating buckets and other aggregates are not countries
        if EXCLUDE_DEFAULTED_FROM_AGGREGATE and ent in DEFAULTED_COUNTRIES:''')

# ---- 6. By_Rating: show the published YTD, with the derived one as fallback
rep('''        ws["A2"] = "Latest snapshot of each JPM rating bucket. YTD TR uses end-of-prior-year as denominator."
        ws.merge_cells("A2:F2")
        self._font(ws["A2"], color=COLOR_NOTE, italic=True)

        header_row = 4
        for i, h in enumerate(["Bucket", "Spread (bps)", "Yield (%)", "Total Return Idx", "YTD TR (%)", "Notes"], start=1):
            self._hdr(ws.cell(row=header_row, column=i), h)
        ws.column_dimensions["A"].width = 22
        ws.column_dimensions["B"].width = 14
        ws.column_dimensions["C"].width = 12
        ws.column_dimensions["D"].width = 18
        ws.column_dimensions["E"].width = 14
        ws.column_dimensions["F"].width = 30''',
'''        ws["A2"] = ("Latest snapshot of each JPM rating bucket. YTD total return is JPM's own "
                    "published figure where the file carries it, decomposed into the spread and "
                    "Treasury legs; otherwise it is derived from the return index against "
                    "end-of-prior-year.")
        ws.merge_cells("A2:H2")
        self._font(ws["A2"], color=COLOR_NOTE, italic=True)

        header_row = 4
        for i, h in enumerate(["Bucket", "Spread (bps)", "Yield (%)", "Total Return Idx",
                               "YTD TR (%)", "of which: spread (%)", "of which: UST (%)",
                               "YTD source"], start=1):
            self._hdr(ws.cell(row=header_row, column=i), h)
        for col, wdt in zip("ABCDEFGH", [22, 14, 12, 18, 14, 20, 18, 26]):
            ws.column_dimensions[col].width = wdt''')

rep('''            tr_start = self.value_at(rating, METRICS["tret"], ytd_idx) if ytd_idx is not None else None
            self._val(ws.cell(row=row, column=2), sp if sp is not None else "", fmt="0", color=COLOR_HARDCODE)
            self._val(ws.cell(row=row, column=3), yd if yd is not None else "", fmt="0.00", color=COLOR_HARDCODE)
            self._val(ws.cell(row=row, column=4), tr if tr is not None else "", fmt="#,##0.00", color=COLOR_HARDCODE)
            if tr is not None and tr_start not in (None, 0):
                self._val(ws.cell(row=row, column=5), (tr / tr_start) - 1, fmt="0.00%;(0.00%);-")
            else:
                self._val(ws.cell(row=row, column=5), "", fmt="0.00%;(0.00%);-")
            ws.cell(row=row, column=6, value="").border = THIN_BORDER
            row += 1''',
'''            tr_start = self.value_at(rating, METRICS["tret"], ytd_idx) if ytd_idx is not None else None
            self._val(ws.cell(row=row, column=2), sp if sp is not None else "", fmt="0", color=COLOR_HARDCODE)
            self._val(ws.cell(row=row, column=3), yd if yd is not None else "", fmt="0.00", color=COLOR_HARDCODE)
            self._val(ws.cell(row=row, column=4), tr if tr is not None else "", fmt="#,##0.00", color=COLOR_HARDCODE)

            # Prefer JPM's published YTD over one derived from the return index:
            # it is their own attribution and it needs only a single snapshot,
            # whereas the derived figure needs an end-of-prior-year data point
            # that a fresh archive will not have.
            pub = self.latest_value(rating, METRICS["tr_ytd"])
            spr = self.latest_value(rating, METRICS["spr_ytd"])
            ust = self.latest_value(rating, METRICS["ust_ytd"])
            pct = "0.00%;(0.00%);-"
            if pub is not None:
                self._val(ws.cell(row=row, column=5), pub / 100.0, fmt=pct, color=COLOR_HARDCODE)
                src_note = "JPM published (snapshot)"
            elif tr is not None and tr_start not in (None, 0):
                self._val(ws.cell(row=row, column=5), (tr / tr_start) - 1, fmt=pct)
                src_note = "derived from return index"
            else:
                self._val(ws.cell(row=row, column=5), "", fmt=pct)
                src_note = "no YTD data"
            self._val(ws.cell(row=row, column=6),
                      spr / 100.0 if spr is not None else "", fmt=pct, color=COLOR_HARDCODE)
            self._val(ws.cell(row=row, column=7),
                      ust / 100.0 if ust is not None else "", fmt=pct, color=COLOR_HARDCODE)
            c = ws.cell(row=row, column=8, value=src_note)
            c.border = THIN_BORDER
            self._font(c, color=COLOR_NOTE, italic=True)
            row += 1

        if not self.ratings_present:
            ws.cell(row=row, column=1,
                    value=("No rating buckets found in any loaded file. The JPM regional export "
                           "names them 'Investment Grade' / 'BB' / 'NR'; the returns file uses "
                           "'Credit IG only' / 'Credit BB only'. Both are mapped — if this row "
                           "still appears, neither file carried a credit-bucket section."))
            self._font(ws.cell(row=row, column=1), color=COLOR_NOTE, italic=True)''')

open(P,"w",encoding="utf-8").write(src)
print(f"OK - applied 6 edits to {P}")
