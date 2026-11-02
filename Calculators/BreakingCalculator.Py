#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Prepared Scrap Transfer Log → Detailed CSV (NO external libraries)
------------------------------------------------------------------
- Reads all .xlsx and .csv transfer logs in this folder
- Extracts valid prepared-scrap transfer rows
- Writes ONE file: BreakingTotals.csv
- Preserves NET (lbs) and NET TONS (computes tons from NET if absent)
- Appends two totals rows:
    1) TOTAL LBS        -> sum of NET (lbs)  [column C]
    2) TOTAL NET TONS   -> sum of NET TONS   [column D]

Output columns:
    DATE, COMMODITY, NET, NET TONS, FROM PILE #, TO PILE #
"""

import csv
import os
from pathlib import Path
from datetime import datetime, timedelta
import zipfile
import xml.etree.ElementTree as ET
import posixpath

OUTPUT_FILE = "BreakingTotals.csv"

# ===== Output columns (fixed order) =====
OUTPUT_COLS = [
    "DATE", "COMMODITY", "NET", "NET TONS", "FROM PILE #", "TO PILE #"
]

# Header normalization map (to interpret sheet headers reliably)
HEADER_ALIASES = {
    "date": "DATE",
    "commodity": "COMMODITY",
    "vehicle": "VEHICLE",
    "gross": "GROSS",
    "tare": "TARE",
    "net": "NET",
    "net tons": "NET TONS",
    "net_tons": "NET TONS",
    "from pile #": "FROM PILE #",
    "from pile # ": "FROM PILE #",  # trailing space variant
    "to pile #": "TO PILE #",
    "signature": "SIGNATURE",
}

NUMERIC_FIELDS = {"GROSS", "TARE", "NET", "NET TONS"}

NS = {
    "main": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
    "r":    "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "rels": "http://schemas.openxmlformats.org/package/2006/relationships",
}

# -------------------- Helpers --------------------

def clean_header(name: str) -> str:
    if name is None:
        return ""
    name = " ".join(str(name).strip().split())
    return HEADER_ALIASES.get(name.lower(), name)

def excel_serial_to_date(n):
    """Excel 1900 date system (base 1899-12-30)."""
    try:
        days = float(n)
        base = datetime(1899, 12, 30)
        return (base + timedelta(days=days)).date()
    except Exception:
        return None

def parse_date_cell(val):
    """Parse date from Excel serial or string formats."""
    if val is None or val == "":
        return None
    try:
        if isinstance(val, (int, float)):
            d = excel_serial_to_date(val)
            if d:
                return d
        if isinstance(val, str) and val.replace(".", "", 1).isdigit():
            d = excel_serial_to_date(float(val))
            if d:
                return d
    except Exception:
        pass
    for fmt in ("%m/%d/%Y", "%Y-%m-%d", "%m/%d/%y"):
        try:
            return datetime.strptime(str(val).strip(), fmt).date()
        except Exception:
            continue
    try:
        return datetime.fromisoformat(str(val).strip()).date()
    except Exception:
        return None

def to_float_or_zero(x):
    try:
        return float(x)
    except Exception:
        return 0.0

def to_clean_str(v):
    """Stringify while removing trailing .0 for whole integers."""
    if isinstance(v, float):
        if v.is_integer():
            return str(int(v))
        return str(v)
    if isinstance(v, int):
        return str(v)
    return ("" if v is None else str(v).strip())

# -------------- XLSX parsing (OpenXML) --------------

def col_letter_to_index(col_ref: str) -> int:
    """Convert A..Z, AA.. -> 0-based index."""
    col = 0
    for ch in col_ref:
        if not ch.isalpha():
            break
        col = col * 26 + (ord(ch.upper()) - ord('A') + 1)
    return col - 1

def cell_ref_to_rc(ref: str):
    """'C12' -> (row_index_zero_based, col_index_zero_based)."""
    letters = "".join([c for c in ref if c.isalpha()])
    digits  = "".join([c for c in ref if c.isdigit()])
    r = int(digits) - 1 if digits else 0
    c = col_letter_to_index(letters) if letters else 0
    return r, c

def load_shared_strings(zf: zipfile.ZipFile):
    """Return list of shared strings if present."""
    try:
        with zf.open("xl/sharedStrings.xml") as f:
            tree = ET.parse(f)
        sst = []
        for si in tree.getroot().findall("main:si", NS):
            txt_parts = []
            for tnode in si.findall(".//main:t", NS):
                txt_parts.append(tnode.text if tnode.text else "")
            sst.append("".join(txt_parts))
        return sst
    except KeyError:
        return []
    except Exception:
        return []

def first_sheet_path(zf: zipfile.ZipFile) -> str:
    """Resolve the first worksheet path using workbook relationships."""
    try:
        with zf.open("xl/workbook.xml") as f:
            wb = ET.parse(f)
        sheets = wb.getroot().findall("main:sheets/main:sheet", NS)
        if not sheets:
            return "xl/worksheets/sheet1.xml"
        first = sheets[0]
        r_id = first.attrib.get(f"{{{NS['r']}}}id")
        with zf.open("xl/_rels/workbook.xml.rels") as f:
            rels = ET.parse(f)
        for rel in rels.getroot().findall("rels:Relationship", NS):
            if rel.attrib.get("Id") == r_id:
                target = rel.attrib.get("Target")
                # Normalize possible absolute/relative targets
                if target.startswith("/"):
                    target = target[1:]
                if not target.startswith("xl/"):
                    target = posixpath.normpath(posixpath.join("xl", target))
                else:
                    target = posixpath.normpath(target)
                return target
        return "xl/worksheets/sheet1.xml"
    except Exception:
        return "xl/worksheets/sheet1.xml"

def read_xlsx_matrix(path: Path):
    """Return first-sheet values as a 2D list of strings."""
    with zipfile.ZipFile(path, "r") as zf:
        sst = load_shared_strings(zf)
        sheet_rel_path = first_sheet_path(zf)
        with zf.open(sheet_rel_path) as f:
            tree = ET.parse(f)
        root = tree.getroot()

        cells = {}
        max_r = max_c = -1
        sheetData = root.find("main:sheetData", NS)
        if sheetData is None:
            return []

        for row in sheetData.findall("main:row", NS):
            for c in row.findall("main:c", NS):
                r_attr = c.attrib.get("r")
                if not r_attr:
                    continue
                r_idx, c_idx = cell_ref_to_rc(r_attr)
                if r_idx > max_r: max_r = r_idx
                if c_idx > max_c: max_c = c_idx

                t = c.attrib.get("t")
                v = ""
                if t == "inlineStr":
                    is_node = c.find("main:is/main:t", NS)
                    v = is_node.text if is_node is not None else ""
                elif t == "s":
                    v_node = c.find("main:v", NS)
                    if v_node is not None and v_node.text is not None:
                        try:
                            s_idx = int(v_node.text)
                            v = sst[s_idx] if 0 <= s_idx < len(sst) else ""
                        except Exception:
                            v = ""
                elif t == "b":
                    v_node = c.find("main:v", NS)
                    v = "TRUE" if (v_node is not None and v_node.text == "1") else "FALSE"
                else:
                    v_node = c.find("main:v", NS)
                    v = v_node.text if v_node is not None else ""

                cells[(r_idx, c_idx)] = v

        rows = max_r + 1
        cols = max_c + 1
        matrix = []
        for r in range(rows):
            matrix.append([cells.get((r, c), "") for c in range(cols)])
        return matrix

# -------------- CSV parsing (built-in) --------------

def read_csv_matrix(path: Path):
    matrix = []
    with open(path, newline="", encoding="utf-8-sig") as f:
        rdr = csv.reader(f)
        for row in rdr:
            matrix.append([cell for cell in row])
    return matrix

# -------------- Build records ----------------------

def detect_header_row(matrix):
    """Find row containing both 'date' and 'commodity'."""
    for i, row in enumerate(matrix):
        lowered = [str(x).strip().lower() for x in row if str(x).strip() != ""]
        if "date" in lowered and "commodity" in lowered:
            return i
    return None

def build_rows_from_matrix(matrix, source_name):
    """
    Convert a matrix into standardized records:
      DATE, COMMODITY, NET, NET TONS, FROM PILE #, TO PILE #
    Drops zero/blank rows; computes NET TONS if missing and NET>0.
    """
    header_row = detect_header_row(matrix)
    if header_row is None:
        return []

    headers = [clean_header(h) for h in matrix[header_row]]
    idx = {name: i for i, name in enumerate(headers) if name}

    def get(row_vals, name):
        j = idx.get(name)
        return row_vals[j] if j is not None and j < len(row_vals) else ""

    results = []
    for r in range(header_row + 1, len(matrix)):
        row_vals = matrix[r]
        if not any(str(x).strip() for x in row_vals):
            continue

        date_raw = get(row_vals, "DATE")
        comm     = get(row_vals, "COMMODITY")
        if not str(date_raw).strip() and not str(comm).strip():
            continue

        d = parse_date_cell(date_raw)
        if d is None:
            continue

        net      = to_float_or_zero(get(row_vals, "NET"))
        net_tons = to_float_or_zero(get(row_vals, "NET TONS"))
        from_p   = to_clean_str(get(row_vals, "FROM PILE #"))
        to_p     = to_clean_str(get(row_vals, "TO PILE #"))

        # Drop zero/blank rows
        if net <= 0:
            continue

        # Compute tons if missing/zero
        if net_tons == 0 and net > 0:
            net_tons = round(net / 2000.0, 3)

        results.append({
            "DATE": str(d),
            "COMMODITY": to_clean_str(comm),
            "NET": net,
            "NET TONS": net_tons,
            "FROM PILE #": from_p,
            "TO PILE #": to_p,
        })

    return results

# -------------- Main -------------------------------

def main():
    base = Path(os.path.abspath(os.path.dirname(__file__)))

    candidates = [
        p for p in base.iterdir()
        if p.is_file()
        and p.suffix.lower() in {".xlsx", ".csv"}
        and not p.name.startswith("~$")
        and p.name != OUTPUT_FILE
    ]

    if not candidates:
        return  # nothing to do

    all_rows = []

    for path in sorted(candidates):
        try:
            if path.suffix.lower() == ".xlsx":
                matrix = read_xlsx_matrix(path)
            else:
                matrix = read_csv_matrix(path)

            rows = build_rows_from_matrix(matrix, path.name)
            if rows:
                all_rows.extend(rows)
        except Exception:
            # ignore single-file failures and continue on others
            continue

    if not all_rows:
        return  # nothing parsed

    # ===== Totals =====
    total_net_lbs  = 0.0    # column C
    total_net_tons = 0.0    # column D
    for r in all_rows:
        try:
            total_net_lbs  += float(r.get("NET", 0) or 0)
        except Exception:
            pass
        try:
            total_net_tons += float(r.get("NET TONS", 0) or 0)
        except Exception:
            pass

    total_net_lbs  = round(total_net_lbs, 3)
    total_net_tons = round(total_net_tons, 3)

    # ===== Write output CSV =====
    with open(base / OUTPUT_FILE, "w", newline="", encoding="utf-8") as out:
        w = csv.writer(out)

        # Header
        w.writerow(OUTPUT_COLS)

        # Detail rows
        for r in all_rows:
            w.writerow([r.get(c, "") for c in OUTPUT_COLS])

        # Totals (two rows)
        # 1) TOTAL LBS in column C (NET)
        w.writerow([
            "",              # DATE
            "TOTAL LBS",     # COMMODITY label
            total_net_lbs,   # NET (lbs)
            "",              # NET TONS (empty)
            "",              # FROM
            ""               # TO
        ])

        # 2) TOTAL NET TONS in column D (NET TONS)
        w.writerow([
            "",                  # DATE
            "TOTAL NET TONS",    # COMMODITY label
            "",                  # NET lbs empty
            total_net_tons,      # NET TONS total
            "",                  # FROM
            ""                   # TO
        ])

if __name__ == "__main__":
    main()