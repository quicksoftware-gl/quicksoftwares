#!/usr/bin/env python3
"""
Regenerate coupons-manifest.js by reading Discount.xlsx.

Run this after editing Discount.xlsx:
    ./update-coupons.py

Discount.xlsx has TWO sheets:

1. "Sheet1" — coupon codes. Five columns:
       Discount_Coupon   Discount_Percentage   Internal_Discount_Percentage   Applies_To   Validity
       50OFF             0.5                   (blank)                        both         2026-12-31
       20OFF             0.2                   (blank)                        windows      (blank)
       BHOLEKRIPA        (blank)               0.25                           macbook      2026-08-15
   Row 1 is the header. Codes are uppercased. Each row should fill exactly
   one of Discount_Percentage (0.0–1.0) or Internal_Discount_Percentage
   (0.0–1.0). Both behave identically as percentage discounts; coupons
   defined via Internal_Discount_Percentage are marked internal-only and
   are hidden from the public "Offers" modal on the site.
   Applies_To controls which service the coupon discounts: "windows",
   "macbook", or "both" (default if blank or unrecognized).
   Validity is the last date the coupon works (inclusive), as a date cell
   or ISO string (YYYY-MM-DD). Blank = no expiry.

2. "Tiers" — the quantity discount shown next to the Windows/Macbook tabs.
   One row per platform, eight columns:
       Platform  Off_Qty2  Off_Qty3  Off_Qty4  Cap_Qty  Charge_Top_N  Cap_Off  Cap_Price
       windows   500       1000      1500      5        5             1500     6000
       macbook   1000      2500      4000      5        5             5000     7500
   Off_Qty2/3/4 = flat ₹ off once that many softwares are in the cart.
   At Cap_Qty or more softwares the customer pays for only the Charge_Top_N
   most-expensive softwares minus Cap_Off; the tab pill then reads
   "<Cap_Qty> or more Softwares at <Cap_Price>". All amounts are stored in
   ₹ as numbers so the frontend can convert them to the active currency at
   the live FX rate. If this sheet is missing, the frontend falls back to
   its built-in defaults, so older Discount.xlsx files keep working.

Both sheets are located by name, so the physical sheetN.xml ordering does
not matter. Output shape:
    window.COUPONS = { ... };
    window.TIERS   = { "win-soft": {...}, "mac-soft": {...} };
"""
import json
import os
import sys
import xml.etree.ElementTree as ET
import zipfile
from datetime import datetime, timedelta

NS = {"main": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}
REL_ID = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id"

script_dir = os.path.dirname(os.path.abspath(__file__))
os.chdir(script_dir)

XLSX = "Discount.xlsx"
OUT = "coupons-manifest.js"

PLATFORM_SVC = {"windows": "win-soft", "macbook": "mac-soft"}


def parse_validity(raw):
    """Convert an Excel date-cell raw value into an ISO date string (YYYY-MM-DD).

    Excel stores dates as serial numbers (days since 1899-12-30 in Windows
    mode). If the user typed an ISO string in a text cell, accept that too.
    Returns None if the value is blank or unparseable.
    """
    if raw is None or raw == "":
        return None
    s = str(raw).strip()
    if not s:
        return None
    # Numeric → Excel serial date.
    try:
        n = float(s)
        if n > 0:
            # 1899-12-30 epoch handles the 1900 leap-year bug for serials > 60.
            dt = datetime(1899, 12, 30) + timedelta(days=int(n))
            return dt.strftime("%Y-%m-%d")
    except (ValueError, TypeError):
        pass
    # ISO string fallback (accepts "2026-12-31" or "2026-12-31T..").
    try:
        return datetime.fromisoformat(s[:10]).strftime("%Y-%m-%d")
    except (ValueError, TypeError):
        return None


def load_shared_strings(z):
    strings = []
    try:
        with z.open("xl/sharedStrings.xml") as f:
            tree = ET.parse(f)
            for si in tree.getroot().findall("main:si", NS):
                # A shared string may be split into several <r><t> runs; join them.
                texts = [t.text or "" for t in si.findall(".//main:t", NS)]
                strings.append("".join(texts))
    except KeyError:
        pass
    return strings


def cell_value(c, strings):
    typ = c.get("t")
    if typ == "inlineStr":
        t = c.find("main:is/main:t", NS)
        return t.text if (t is not None and t.text is not None) else ""
    v = c.find("main:v", NS)
    raw = v.text if v is not None else ""
    if typ == "s" and raw != "":
        return strings[int(raw)]
    return raw


def sheet_paths(z):
    """Map worksheet display name -> zip path, in workbook order."""
    rels = {}
    with z.open("xl/_rels/workbook.xml.rels") as f:
        for rel in ET.parse(f).getroot():
            rels[rel.get("Id")] = rel.get("Target")
    out = {}
    with z.open("xl/workbook.xml") as f:
        for sh in ET.parse(f).getroot().findall(".//main:sheet", NS):
            target = rels.get(sh.get(REL_ID), "")
            if not target:
                continue
            t = target.lstrip("/")
            if not t.startswith("xl/"):
                t = "xl/" + t
            out[sh.get("name")] = t
    return out


def read_rows(z, path, strings):
    """Return a list of {column_letter: value} dicts, one per <row>."""
    rows = []
    with z.open(path) as f:
        for row in ET.parse(f).getroot().findall(".//main:row", NS):
            cells = {}
            for c in row.findall("main:c", NS):
                ref = c.get("r", "")  # e.g. "A2"
                col = "".join(ch for ch in ref if ch.isalpha())
                cells[col] = cell_value(c, strings)
            rows.append(cells)
    return rows


def as_number(v):
    if v in (None, ""):
        return None
    try:
        n = float(v)
    except (ValueError, TypeError):
        return None
    return int(n) if n.is_integer() else n


def parse_coupons(rows):
    coupons = {}
    header_seen = False
    for row_cells in rows:
        if not row_cells.get("A"):
            continue
        if not header_seen:
            header_seen = True
            continue
        code = str(row_cells["A"]).strip().upper()
        if not code:
            continue
        pct_raw = row_cells.get("B", "")
        internal_raw = row_cells.get("C", "")
        applies_raw = (row_cells.get("D", "") or "").strip().lower()
        validity = parse_validity(row_cells.get("E", ""))
        pct = None
        internal_pct = None
        try:
            if pct_raw not in (None, ""):
                pct = float(pct_raw)
        except (ValueError, TypeError):
            print(f"  ! {code}: invalid Discount_Percentage '{pct_raw}' — ignored", file=sys.stderr)
        try:
            if internal_raw not in (None, ""):
                internal_pct = float(internal_raw)
        except (ValueError, TypeError):
            print(f"  ! {code}: invalid Internal_Discount_Percentage '{internal_raw}' — ignored", file=sys.stderr)

        if applies_raw in ("windows", "macbook", "both"):
            applies_to = applies_raw
        else:
            if applies_raw:
                print(f"  ! {code}: unknown Applies_To '{applies_raw}', defaulting to 'both'", file=sys.stderr)
            applies_to = "both"

        if internal_pct is not None and 0 <= internal_pct <= 1:
            coupons[code] = {
                "type": "percent",
                "value": internal_pct,
                "applies_to": applies_to,
                "internal": True,
            }
        elif pct is not None and 0 <= pct <= 1:
            coupons[code] = {"type": "percent", "value": pct, "applies_to": applies_to}
        else:
            print(f"  ! Skipping {code}: needs Discount_Percentage (0..1) OR Internal_Discount_Percentage (0..1)", file=sys.stderr)
            continue
        if validity:
            coupons[code]["validity"] = validity
    return coupons


def parse_tiers(rows):
    """Parse the Tiers sheet into { svc_id: {off, capQty, topN, capOff, capPrice} }."""
    tiers = {}
    header_seen = False
    for cells in rows:
        plat = (cells.get("A", "") or "").strip().lower()
        if not plat:
            continue
        if not header_seen:
            header_seen = True
            continue
        svc_id = PLATFORM_SVC.get(plat)
        if not svc_id:
            print(f"  ! Tiers: unknown Platform '{plat}' — skipped", file=sys.stderr)
            continue
        off = {}
        for col, qty in (("B", "2"), ("C", "3"), ("D", "4")):
            n = as_number(cells.get(col, ""))
            if n is not None:
                off[qty] = n
        cap_qty = as_number(cells.get("E", ""))
        top_n = as_number(cells.get("F", ""))
        cap_off = as_number(cells.get("G", ""))
        cap_price = as_number(cells.get("H", ""))
        tiers[svc_id] = {
            "off": off,
            "capQty": int(cap_qty) if cap_qty else 5,
            "topN": int(top_n) if top_n else 5,
            "capOff": cap_off if cap_off is not None else 0,
            "capPrice": cap_price if cap_price is not None else 0,
        }
    return tiers


def read_xlsx(path):
    if not os.path.exists(path):
        return {}, {}
    with zipfile.ZipFile(path) as z:
        strings = load_shared_strings(z)
        paths = sheet_paths(z)
        names = list(paths.keys())

        # Coupons: the sheet literally named "Sheet1", else the first sheet.
        coupon_sheet = "Sheet1" if "Sheet1" in paths else (names[0] if names else None)
        coupons = parse_coupons(read_rows(z, paths[coupon_sheet], strings)) if coupon_sheet else {}

        # Tiers: the sheet named "Tiers" (case-insensitive), if present.
        tier_name = next((n for n in names if n.strip().lower() == "tiers"), None)
        tiers = parse_tiers(read_rows(z, paths[tier_name], strings)) if tier_name else {}

    return coupons, tiers


def main():
    coupons, tiers = read_xlsx(XLSX)
    with open(OUT, "w", encoding="utf-8", newline="\n") as f:
        f.write("// Auto-generated by update-coupons.py - do not edit by hand.\n")
        f.write(f"// Source: {XLSX}. Run ./update-coupons.py after editing.\n")
        f.write("window.COUPONS = " + json.dumps(coupons, indent=2) + ";\n")
        f.write("window.TIERS = " + json.dumps(tiers, indent=2) + ";\n")
    print(f"Wrote {len(coupons)} coupon(s) and {len(tiers)} tier config(s) to {OUT}")


if __name__ == "__main__":
    main()
