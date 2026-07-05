#!/usr/bin/env python3
"""
Regenerate softwares-manifest.js from Softwares.xlsx.

Reads the in-repo Excel file Softwares.xlsx (override via SOFTWARES_XLSX
env var). The workbook has two sheets — "Windows" and "Macbook". Row 1
is the header. Column A holds the software name, column B holds an
optional Version label (text). Softwares whose Version equals "Latest"
(case-insensitive) trigger a platform-specific surcharge in the UI
(+₹500 for Windows, +₹1500 for Macbook) and show a "Price Changed due
to the Latest Version" note.

Outputs softwares-manifest.js as:

    window.SOFTWARES = {
      "windows": [
        { "name": "Autocad 2026", "row": 2, "version": "Latest" },
        { "name": "Autocad 2024", "row": 3 },
        ...
      ],
      "macbook": [ ... ]
    };

The `row` field is the 1-based Excel row number and is used by the
Software Catalog UI in index.html to load the matching product image
from Windows_Pic/<row>.{jpg,png,webp} or Macbook_Pic/<row>.{jpg,png,webp}.

Run after editing the sheet:
    ./update-softwares.py
"""
import json
import os
import sys
import xml.etree.ElementTree as ET
import zipfile

NS = {"main": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}
REL_NS = {"rel": "http://schemas.openxmlformats.org/package/2006/relationships"}

script_dir = os.path.dirname(os.path.abspath(__file__))
os.chdir(script_dir)

XLSX = os.environ.get("SOFTWARES_XLSX", "Softwares.xlsx")
OUT = "softwares-manifest.js"


def load_shared_strings(z):
    strings = []
    try:
        with z.open("xl/sharedStrings.xml") as f:
            tree = ET.parse(f)
            for si in tree.getroot().findall("main:si", NS):
                # A shared string can contain multiple <t> runs; concatenate them.
                parts = [t.text or "" for t in si.findall(".//main:t", NS)]
                strings.append("".join(parts))
    except KeyError:
        pass
    return strings


def sheet_name_to_path(z):
    """Return { sheet_name: zip_path_to_worksheet_xml }."""
    with z.open("xl/workbook.xml") as f:
        wb = ET.parse(f).getroot()
    with z.open("xl/_rels/workbook.xml.rels") as f:
        rels = ET.parse(f).getroot()

    rel_target = {}
    for rel in rels.findall("rel:Relationship", REL_NS):
        rel_target[rel.get("Id")] = rel.get("Target")

    mapping = {}
    for sh in wb.findall("main:sheets/main:sheet", NS):
        name = sh.get("name")
        rid = sh.get("{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id")
        target = rel_target.get(rid, "")
        # Targets are relative to xl/, e.g. "worksheets/sheet1.xml".
        if target.startswith("/"):
            zip_path = target.lstrip("/")
        else:
            zip_path = "xl/" + target
        mapping[name] = zip_path
    return mapping


def cell_value(c, strings):
    typ = c.get("t")
    if typ == "inlineStr":
        t = c.find("main:is/main:t", NS)
        return t.text if (t is not None and t.text) else ""
    v = c.find("main:v", NS)
    raw = v.text if v is not None else ""
    if typ == "s" and raw != "":
        try:
            return strings[int(raw)]
        except (ValueError, IndexError):
            return ""
    return raw or ""


def read_sheet_rows(z, zip_path, strings):
    """Return [{name, row, version?}] for each data row.
    Column A = software name (required). Column B = version (optional).
    `row` is the 1-based Excel row number — used by the catalog UI to look
    up the matching image in Windows_Pic/<row>.jpg or Macbook_Pic/<row>.jpg."""
    with z.open(zip_path) as f:
        tree = ET.parse(f)
    items = []
    seen = set()
    header_seen = False
    for row in tree.getroot().findall(".//main:row", NS):
        cells = {}
        for c in row.findall("main:c", NS):
            ref = c.get("r", "")
            col = "".join(ch for ch in ref if ch.isalpha())
            if col in ("A", "B"):
                cells[col] = c
        a_cell = cells.get("A")
        if a_cell is None:
            continue
        name = cell_value(a_cell, strings).strip()
        if not name:
            continue
        if not header_seen:
            header_seen = True
            continue
        key = name.lower()
        if key in seen:
            continue
        seen.add(key)
        try:
            row_num = int(row.get("r") or 0)
        except ValueError:
            row_num = 0
        item = {"name": name, "row": row_num}
        b_cell = cells.get("B")
        if b_cell is not None:
            version = cell_value(b_cell, strings).strip()
            if version:
                item["version"] = version
        items.append(item)
    return items


def read_softwares(path):
    if not os.path.exists(path):
        print(f"  ! Source not found: {path}", file=sys.stderr)
        return {"windows": [], "macbook": []}

    out = {"windows": [], "macbook": []}
    with zipfile.ZipFile(path) as z:
        strings = load_shared_strings(z)
        sheets = sheet_name_to_path(z)
        for sheet_name, zip_path in sheets.items():
            key = sheet_name.strip().lower()
            if key not in out:
                # Ignore unexpected sheets but log them.
                print(f"  ! Skipping sheet '{sheet_name}' (expected Windows or Macbook)", file=sys.stderr)
                continue
            out[key] = read_sheet_rows(z, zip_path, strings)
    return out


def main():
    if not os.path.exists(XLSX):
        # Refuse to overwrite a good manifest with empties (e.g. when cron
        # can't see ~/Downloads because macOS Full Disk Access isn't granted).
        print(f"  ! Source not found: {XLSX} — leaving {OUT} untouched", file=sys.stderr)
        sys.exit(0)
    data = read_softwares(XLSX)
    if not data["windows"] and not data["macbook"]:
        print(f"  ! Parsed 0 entries from {XLSX} — leaving {OUT} untouched", file=sys.stderr)
        sys.exit(0)
    payload = json.dumps(data, indent=2, ensure_ascii=False)
    with open(OUT, "w", encoding="utf-8") as f:
        f.write("// Auto-generated by update-softwares.py — do not edit by hand.\n")
        f.write(f"// Source: {XLSX}. Run ./update-softwares.py after editing.\n")
        f.write("window.SOFTWARES = " + payload + ";\n")
    print(f"✓ Wrote {len(data['windows'])} windows + {len(data['macbook'])} macbook software(s) to {OUT}")


if __name__ == "__main__":
    main()
