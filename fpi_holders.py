"""
Nifty 500 stocks ka FPI holders count (Nos. of shareholders) -> Excel.

Source: NSE Shareholding Pattern XBRL filings (official, free).

Usage:
    python fpi_holders.py                          # all Nifty 500 -> fpi_holders.xlsx
    python fpi_holders.py --symbols HDFCBANK,TCS   # only these
    python fpi_holders.py --symbols HDFCBANK --debug
"""
import argparse
import csv
import io
import re
import time
import xml.etree.ElementTree as ET
from datetime import datetime

import requests
from openpyxl import Workbook
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter

BASE = "https://www.nseindia.com"
NIFTY500 = "https://archives.nseindia.com/content/indices/ind_nifty500list.csv"
HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
                  "(KHTML, like Gecko) Chrome/124.0 Safari/537.36",
    "Accept": "application/json, text/plain, */*",
    "Accept-Language": "en-US,en;q=0.9",
    "Referer": BASE + "/companies-listing/corporate-filings-shareholding-pattern",
}


def session():
    s = requests.Session()
    s.headers.update(HEADERS)
    s.get(BASE, timeout=30)
    s.get(BASE + "/companies-listing/corporate-filings-shareholding-pattern", timeout=30)
    return s


def get_universe(s):
    """{symbol: company name} for Nifty 500."""
    r = s.get(NIFTY500, timeout=30)
    r.raise_for_status()
    return {row["Symbol"].strip(): row["Company Name"].strip()
            for row in csv.DictReader(io.StringIO(r.text))}


def latest_filing(s, symbol):
    r = s.get(f"{BASE}/api/corporate-share-holdings-master",
              params={"index": "equities", "symbol": symbol}, timeout=30)
    r.raise_for_status()
    data = r.json()
    if isinstance(data, dict):
        data = data.get("data", [])
    data = [d for d in data if d.get("xbrl")]
    if not data:
        return None

    def key(d):
        for f in ("%d-%b-%Y", "%d-%m-%Y", "%Y-%m-%d"):
            try:
                return datetime.strptime(d.get("date", ""), f)
            except ValueError:
                pass
        return datetime.min

    return max(data, key=key)


def local(tag):
    return tag.rsplit("}", 1)[-1]


def fpi_holders(xml_bytes, debug=False):
    root = ET.fromstring(xml_bytes)
    ctx = {}
    for c in root.iter():
        if local(c.tag) == "context":
            ctx[c.attrib.get("id")] = " ".join(
                (m.text or "") for m in c.iter() if local(m.tag) == "explicitMember")
    total = 0
    for e in root.iter():
        name = local(e.tag).lower()
        ref = e.attrib.get("contextRef")
        if ref is None or not e.text:
            continue
        if "foreignportfolio" not in ctx.get(ref, "").lower():
            continue
        if "numberofshareholders" in name or "nosofshareholders" in name:
            try:
                val = int(float(e.text.strip()))
            except ValueError:
                continue
            if debug:
                print("HOLDERS", local(e.tag), ctx[ref], val)
            total += val
    return total


def save_xlsx(rows, path):
    wb = Workbook()
    ws = wb.active
    ws.title = "FPI Holders"
    ws.append(["Stock Name", "Symbol", "No. of FPI Holders", "Quarter"])
    for c in ws[1]:
        c.font = Font(bold=True)
    for r in sorted(rows, key=lambda x: x[2], reverse=True):
        ws.append(r)
    for i, w in enumerate([38, 14, 20, 14], 1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.freeze_panes = "A2"
    wb.save(path)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--symbols")
    ap.add_argument("--out", default="fpi_holders.xlsx")
    ap.add_argument("--debug", action="store_true")
    a = ap.parse_args()

    s = session()
    universe = get_universe(s)
    if a.symbols:
        syms = [x.strip().upper() for x in a.symbols.split(",")]
        universe = {k: universe.get(k, k) for k in syms}

    rows = []
    for i, (sym, name) in enumerate(universe.items(), 1):
        try:
            rec = latest_filing(s, sym)
            if not rec:
                print(f"[{i}/{len(universe)}] {sym}: no filing")
                continue
            r = s.get(rec["xbrl"], timeout=60)
            r.raise_for_status()
            n = fpi_holders(r.content, a.debug)
            rows.append([name, sym, n, rec.get("date")])
            print(f"[{i}/{len(universe)}] {sym}: {n}")
        except Exception as ex:
            print(f"[{i}/{len(universe)}] {sym}: ERROR {ex}")
            if "401" in str(ex) or "403" in str(ex):
                s = session()
        time.sleep(1.5)

    save_xlsx(rows, a.out)
    print(f"Saved {len(rows)} rows -> {a.out}")


if __name__ == "__main__":
    main()
