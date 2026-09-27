#!/usr/bin/env python3
"""Embed skat.dk's ABIS list (aktiebaserede investeringsselskaber) into index.html.

Usage: python3 build-abis.py [path/to/abis-liste.xlsx] [--label "august 2026"]

Reads every sheet like ingestAbis() in the app: header row = first row with a cell containing "isin";
columns ISIN, "Registrerede år" (list of years, current skat.dk layout) or, for older layouts,
"Første registreringsår" + "Fjernet". Writes `const ABIS_EMBED = {...}` between
/* ABIS_EMBED_START */ and /* ABIS_EMBED_END */ in the HTML next to this script.
Groups: [spec, "ISIN,ISIN,..."], spec = registered years "2021,2022,2026" (union over all rows of the
ISIN) or "first-removed" for the older layout. An ISIN is on the list for year T if T is in its years.
"""
import datetime, json, os, re, sys
from collections import defaultdict
import openpyxl

HERE = os.path.dirname(os.path.abspath(__file__))
HTML = os.path.join(HERE, "index.html")
ISIN = re.compile(r"\b([A-Z]{2}[A-Z0-9]{9}\d)\b")
YEAR = re.compile(r"(?<!\d)(20\d\d)(?!\d)")

def years(v):
    if v is None: return set()
    if isinstance(v, (datetime.date, datetime.datetime)): return {v.year}
    return {int(y) for y in YEAR.findall(str(v))}

def year_of(v):
    if v is None or str(v).strip() == "": return None
    if isinstance(v, (datetime.date, datetime.datetime)): return v.year
    m = re.search(r"(\d{4})\s*$", str(v).strip())
    return int(m.group(1)) if m else None

def main():
    args = [a for a in sys.argv[1:]]
    label = None
    if "--label" in args:
        i = args.index("--label"); label = args[i + 1]; del args[i:i + 2]
    src = args[0] if args else os.path.join(HERE, "samples", "august-2026-abis-liste-2021-2026.xlsx")
    if not label:
        m = re.match(r"([a-zæøå]+)-(\d{4})", os.path.basename(src).lower())
        label = f"{m.group(1)} {m.group(2)}" if m else os.path.basename(src)

    wb = openpyxl.load_workbook(src, read_only=True, data_only=True)
    reg = defaultdict(set)      # isin -> registered years (union)
    old = {}                    # isin -> (first, removed) for older layout
    for ws in wb.worksheets:
        rows = ws.iter_rows(values_only=True)
        hdr = None
        for r in rows:
            if any(c is not None and "isin" in str(c).lower() for c in r):
                hdr = [str(c or "").replace("\xa0", " ").strip().lower() for c in r]; break
        if hdr is None: continue
        i_isin = next(i for i, h in enumerate(hdr) if "isin" in h)
        i_reg = next((i for i, h in enumerate(hdr) if re.match(r"(registrerede|registered)", h)), None)
        i_first = next((i for i, h in enumerate(hdr) if re.search(r"første|first", h)), None)
        i_rem = next((i for i, h in enumerate(hdr) if re.search(r"fjernet|removed", h)), None)
        for r in rows:
            if i_isin >= len(r): continue
            m = ISIN.search(str(r[i_isin] or "").upper())
            if not m: continue
            isin = m.group(1)
            if i_reg is not None:
                reg[isin] |= years(r[i_reg] if i_reg < len(r) else None)
            else:
                first = year_of(r[i_first]) if i_first is not None else None
                rem = year_of(r[i_rem]) if i_rem is not None else None
                old[isin] = (first, rem)

    groups = defaultdict(list)
    no_years = 0
    for isin, ys in reg.items():
        if not ys: no_years += 1; continue            # never registered -> never on the list
        groups[",".join(str(y) for y in sorted(ys))].append(isin)
    for isin, (first, rem) in old.items():
        if isin in reg: continue
        groups[f"{first or ''}-{rem or ''}"].append(isin)

    ordered = sorted(groups.items(), key=lambda kv: (-len(kv[1]), kv[0]))
    body = ",\n".join(json.dumps([spec, ",".join(sorted(isins))]) for spec, isins in ordered)
    block = (f"/* ABIS_EMBED_START */\n"
             f"const ABIS_EMBED = {{label:{json.dumps(label, ensure_ascii=False)}, source:{json.dumps(os.path.basename(src), ensure_ascii=False)}, groups:[\n{body}\n]}};\n"
             f"/* ABIS_EMBED_END */")
    html = open(HTML, encoding="utf-8").read()
    new, n = re.subn(r"/\* ABIS_EMBED_START \*/.*?/\* ABIS_EMBED_END \*/", lambda _: block, html, flags=re.S)
    if n != 1: sys.exit("ABIS_EMBED markers not found exactly once in " + HTML)
    open(HTML, "w", encoding="utf-8").write(new)

    total = sum(len(v) for v in groups.values())
    print(f"source: {src}\nlabel: {label}\ngroups: {len(groups)}, ISIN: {total} (+{no_years} listed without any registered year, skipped)")
    for spec, isins in ordered: print(f"  {len(isins):5d}  {spec}")
    allyears = sorted({int(y) for spec in groups if "-" not in spec for y in spec.split(",")})
    for y in allyears:
        print(f"  on list for {y}: {sum(len(v) for s, v in groups.items() if '-' not in s and str(y) in s.split(','))}")
    print(f"embedded block: {len(block)/1024:.1f} KB")

if __name__ == "__main__":
    main()
