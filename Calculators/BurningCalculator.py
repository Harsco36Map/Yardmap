
import zipfile
import xml.etree.ElementTree as ET
import csv
import re
from pathlib import Path
from datetime import date, datetime, timedelta

INPUT_DIR = Path(".")
OUTPUT_CSV = "BurningTotals.csv"

HEADER_SYNONYMS = {
    "DATE": {"DATE", "DAY", "LOG DATE"},
    "COMMODITY": {"COMMODITY"},
    "GROSS": {"GROSS"},
    "TARE": {"TARE"},
    "NET": {"NET"},
    "NET TONS": {"NET TONS", "NET TNS", "NET (TONS)"},
    "FROM PILE #": {"FROM PILE #", "FROM PILE", "FROM PILE NO"},
    "TO PILE #": {"TO PILE #", "TO PILE", "TO PILE NO"},
    "# OF CUTS": {"# OF CUTS", "CUTS"},
    "BILLABLE NTS": {"BILLABLE NTS", "BILLABLE TONS"},
}

# Final headers in desired output order
OUTPUT_HEADERS = [
    "Date",
    "Net(lbs)",
    "From",
    "To",
    "NetTons",
    "# of Cuts",
    "Billable Tons"
]

def col_letter_to_index(col_letters):
    n = 0
    for c in col_letters:
        n = n * 26 + (ord(c.upper()) - 64)
    return n - 1

def parse_cell_value(cel, shared_strings):
    ns = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
    t = cel.get("t")

    if t == "s":  # shared string
        v = cel.find(f"{ns}v")
        return shared_strings[int(v.text)] if v is not None else ""

    if t == "inlineStr":
        is_el = cel.find(f"{ns}is")
        t_el = is_el.find(f"{ns}t") if is_el is not None else None
        return t_el.text if t_el is not None else ""

    v = cel.find(f"{ns}v")
    return v.text if v is not None else ""

def read_shared_strings(zf):
    try:
        with zf.open("xl/sharedStrings.xml") as f:
            root = ET.fromstring(f.read())
        ns = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
        strings = []
        for si in root.findall(f"{ns}si"):
            text = "".join((t.text or "") for t in si.findall(f"{ns}t"))
            strings.append(text)
        return strings
    except:
        return []

def read_first_sheet_xml(zf):
    sheets = sorted(
        n for n in zf.namelist()
        if n.startswith("xl/worksheets/sheet") and n.endswith(".xml")
    )
    with zf.open(sheets[0]) as f:
        return ET.fromstring(f.read())

def excel_serial_to_iso(val):
    """Use correct mapping: date = 1899-12-30 + serial days (no -1 offset)."""
    try:
        serial = int(float(val))
    except:
        return val
    base = date(1899, 12, 30)
    d = base + timedelta(days=serial)
    return d.isoformat()

def find_header_row(rows):
    for i, row in enumerate(rows):
        norm = [str(c).strip().upper() for c in row]
        colmap = {}
        for canon, syns in HEADER_SYNONYMS.items():
            for idx, cell in enumerate(norm):
                if cell in syns:
                    colmap[canon] = idx
                    break
        essentials = {"DATE", "COMMODITY", "NET TONS", "FROM PILE #", "TO PILE #", "BILLABLE NTS"}
        if essentials.issubset(colmap):
            return i, colmap
    return None, {}

def row_is_footer_or_blank(r):
    joined = " ".join(str(v).strip().upper() for v in r if v not in ("", None))
    if joined == "":
        return True
    if "FORM HF-03" in joined or "REV." in joined:
        return True
    nonempty = [str(v).strip() for v in r if v not in ("", None)]
    if nonempty and all(re.fullmatch(r"0+(\\.0+)?", v) for v in nonempty):
        return True
    return False

def parse_xlsx(path):
    with zipfile.ZipFile(path, "r") as z:
        shared = read_shared_strings(z)
        root = read_first_sheet_xml(z)
        ns = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
        rows_xml = root.findall(f".//{ns}row")

        table_rows = []
        for r in rows_xml:
            rmap = {}
            for c in r.findall(f"{ns}c"):
                ref = c.get("r", "")
                m = re.match(r"([A-Z]+)([0-9]+)", ref)
                if not m:
                    continue
                col_letters, _ = m.groups()
                idx = col_letter_to_index(col_letters)
                rmap[idx] = parse_cell_value(c, shared)
            if rmap:
                table_rows.append([rmap.get(i, "") for i in range(max(rmap.keys()) + 1)])
            else:
                table_rows.append([])

    header_idx, col_map = find_header_row(table_rows)
    if header_idx is None:
        return []

    transfers = []
    for r in table_rows[header_idx + 1:]:
        if row_is_footer_or_blank(r):
            break

        date_raw = r[col_map["DATE"]]
        commodity = r[col_map["COMMODITY"]]

        if str(date_raw).strip() == "" and str(commodity).strip() == "":
            break

        entry = {
            "Date": excel_serial_to_iso(date_raw),
            "Net(lbs)": r[col_map.get("NET", "")],
            "From": r[col_map.get("FROM PILE #", "")],
            "To": r[col_map.get("TO PILE #", "")],
            "NetTons": r[col_map["NET TONS"]],
            "# of Cuts": r[col_map.get("# OF CUTS", "")],
            "Billable Tons": r[col_map["BILLABLE NTS"]],
        }
        transfers.append(entry)

    return transfers

def main():
    all_rows = []
    for f in sorted(INPUT_DIR.glob("*.xlsx")):
        all_rows.extend(parse_xlsx(f))

    if not all_rows:
        print("No transfers found.")
        return

    with open(OUTPUT_CSV, "w", newline="", encoding="utf-8") as out:
        writer = csv.DictWriter(out, fieldnames=OUTPUT_HEADERS)
        writer.writeheader()

        for row in all_rows:
            writer.writerow(row)

        last_excel_row = 1 + len(all_rows)  # header = row 1, data starts at row 2

        writer.writerow({})  # blank line

        totals = {h: "" for h in OUTPUT_HEADERS}
        totals["NetTons"] = f"=SUM(E2:E{last_excel_row})"
        totals["Billable Tons"] = f"=SUM(G2:G{last_excel_row})"
        writer.writerow(totals)

    print(f"Created {OUTPUT_CSV}")

if __name__ == "__main__":
    main()
