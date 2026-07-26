#!/usr/bin/env python3
"""
Regenerate socialmedia-manifest.js by reading SocialMedia.xlsx.

Run this after editing SocialMedia.xlsx:
    ./update-socialmedia.py

SocialMedia.xlsx has ONE sheet named "SocialMedia" with a header row and
one data row per currency. Columns (matched by header name, so their order
does not matter):

    Currency   Whatsapp                      Instagram              Youtube
    INR        https://wa.me/917838127423    https://instagram/..   https://youtube/..
    USD        https://wa.me/16472614931     https://instagram/..   https://youtube/..

- Currency is the frontend currency code (INR, USD, ...), uppercased.
- Whatsapp may be a full https://wa.me/<digits> link OR a bare phone number
  (with or without + / spaces); a bare number is normalised to a wa.me link.
- Instagram / Youtube are full profile/channel URLs.

Blank cells are omitted so the frontend falls back to its built-in default
for that platform. If the sheet or file is missing, an empty manifest is
written and the site keeps working on its hardcoded defaults.

Output shape:
    window.SOCIALS = {
      "INR": { "whatsapp": "...", "instagram": "...", "youtube": "..." },
      "USD": { ... }
    };
"""
import json
import os
import re
import sys
import xml.etree.ElementTree as ET
import zipfile

NS = {"main": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}
REL_ID = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id"

script_dir = os.path.dirname(os.path.abspath(__file__))
os.chdir(script_dir)

XLSX = "SocialMedia.xlsx"
OUT = "socialmedia-manifest.js"

# Header name (lowercased) -> manifest key.
PLATFORM_HEADERS = {"whatsapp": "whatsapp", "instagram": "instagram", "youtube": "youtube"}


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


def normalize_whatsapp(raw):
    """A full URL is kept as-is; a bare phone number becomes a wa.me link."""
    s = (raw or "").strip()
    if not s:
        return ""
    if re.match(r"^https?://", s, re.I):
        return s
    digits = re.sub(r"\D", "", s)
    return "https://wa.me/" + digits if digits else s


def parse_socials(rows):
    """Rows -> { CURRENCY: {whatsapp, instagram, youtube} } using header names."""
    socials = {}
    header = None  # {manifest_key_or_"currency": column_letter}
    for cells in rows:
        if header is None:
            # First non-empty row is the header; locate columns by name.
            labels = {col: str(val).strip().lower() for col, val in cells.items() if str(val).strip()}
            if not labels:
                continue
            header = {}
            for col, label in labels.items():
                if label == "currency":
                    header["currency"] = col
                elif label in PLATFORM_HEADERS:
                    header[PLATFORM_HEADERS[label]] = col
            if "currency" not in header:
                print("  ! No 'Currency' column found in SocialMedia sheet.", file=sys.stderr)
                return {}
            continue

        currency = str(cells.get(header["currency"], "") or "").strip().upper()
        if not currency:
            continue
        entry = {}
        for key in ("whatsapp", "instagram", "youtube"):
            col = header.get(key)
            if not col:
                continue
            val = str(cells.get(col, "") or "").strip()
            if not val:
                continue
            entry[key] = normalize_whatsapp(val) if key == "whatsapp" else val
        if entry:
            socials[currency] = entry
    return socials


def read_xlsx(path):
    if not os.path.exists(path):
        return {}
    with zipfile.ZipFile(path) as z:
        strings = load_shared_strings(z)
        paths = sheet_paths(z)
        names = list(paths.keys())
        # The sheet named "SocialMedia" (case-insensitive), else the first sheet.
        sheet = next((n for n in names if n.strip().lower() == "socialmedia"), None)
        if sheet is None:
            sheet = names[0] if names else None
        if sheet is None:
            return {}
        return parse_socials(read_rows(z, paths[sheet], strings))


def main():
    socials = read_xlsx(XLSX)
    with open(OUT, "w", encoding="utf-8", newline="\n") as f:
        f.write("// Auto-generated by update-socialmedia.py - do not edit by hand.\n")
        f.write(f"// Source: {XLSX}. Run ./update-socialmedia.py after editing.\n")
        f.write("window.SOCIALS = " + json.dumps(socials, indent=2) + ";\n")
    print(f"Wrote {len(socials)} currency social-link set(s) to {OUT}")


if __name__ == "__main__":
    main()
