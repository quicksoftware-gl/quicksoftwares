#!/usr/bin/env python3
"""
Regenerate coupons-manifest.js by reading Discount.xlsx.

Run this after editing Discount.xlsx:
    ./update-coupons.py

The .xlsx supports two coupon types via five columns:
    Discount_Coupon   Discount_Percentage   Internal_Discount_Percentage   Applies_To   Validity
    50OFF             0.5                   (blank)                        both         2026-12-31
    20OFF             0.2                   (blank)                        windows      (blank)
    BHOLEKRIPA        (blank)               0.25                           macbook      2026-08-15
    ...

Row 1 is the header. Codes are uppercased. Each row should fill exactly
one of Discount_Percentage (0.0–1.0) or Internal_Discount_Percentage
(0.0–1.0). Both behave identically as percentage discounts; coupons
defined via Internal_Discount_Percentage are marked internal-only and
are hidden from the public "Offers" modal on the site.
Applies_To controls which service the coupon discounts: "windows",
"macbook", or "both" (default if blank or unrecognized).
Validity is the last date the coupon works (inclusive), as a date cell
or ISO string (YYYY-MM-DD). Blank = no expiry.
"""
import json
import os
import sys
import xml.etree.ElementTree as ET
import zipfile
from datetime import datetime, timedelta

NS = {"main": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}

script_dir = os.path.dirname(os.path.abspath(__file__))
os.chdir(script_dir)

XLSX = "Discount.xlsx"
OUT = "coupons-manifest.js"


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


def read_xlsx(path):
    if not os.path.exists(path):
        return {}
    coupons = {}
    with zipfile.ZipFile(path) as z:
        # Shared strings (Excel stores strings in a separate table)
        strings = []
        try:
            with z.open("xl/sharedStrings.xml") as f:
                tree = ET.parse(f)
                for si in tree.getroot().findall("main:si", NS):
                    t = si.find(".//main:t", NS)
                    strings.append(t.text if (t is not None and t.text) else "")
        except KeyError:
            pass

        with z.open("xl/worksheets/sheet1.xml") as f:
            tree = ET.parse(f)
            rows = tree.getroot().findall(".//main:row", NS)
            header_seen = False
            for row in rows:
                # Map cells by column letter so missing cells don't shift indices
                row_cells = {}
                for c in row.findall("main:c", NS):
                    ref = c.get("r", "")  # e.g. "A2"
                    col = "".join(ch for ch in ref if ch.isalpha())
                    typ = c.get("t")
                    val = ""
                    if typ == "inlineStr":
                        # <c><is><t>text</t></is></c>
                        t = c.find("main:is/main:t", NS)
                        if t is not None and t.text is not None:
                            val = t.text
                    else:
                        v = c.find("main:v", NS)
                        raw = v.text if v is not None else ""
                        if typ == "s" and raw != "":
                            val = strings[int(raw)]
                        else:
                            val = raw
                    row_cells[col] = val
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


def main():
    coupons = read_xlsx(XLSX)
    with open(OUT, "w") as f:
        f.write("// Auto-generated by update-coupons.py — do not edit by hand.\n")
        f.write(f"// Source: {XLSX}. Run ./update-coupons.py after editing.\n")
        f.write("window.COUPONS = " + json.dumps(coupons, indent=2) + ";\n")
    print(f"✓ Wrote {len(coupons)} coupon(s) to {OUT}")


if __name__ == "__main__":
    main()
