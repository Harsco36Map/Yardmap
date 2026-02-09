#!/usr/bin/env python3
"""
Tiny converter: OPERS CSV -> stockData.json (includes report_date)

Usage:
  python csv_to_stock_json.py [optional_path_to_csv]

- If no CSV is provided, picks the most recent *.csv in the current folder.
- Writes stockData.json in the same folder.
- Adds a top-level "meta" with {"report_date": "DD-Mon"}, parsed from the CSV header line
  e.g., "OPERS Inventory Report created at 31-DEC-2025 00:00:33" -> "31-Dec".
"""
import csv, json, re, sys
from pathlib import Path

EXPECTED = [
    "Code","Material","Pile","Month Begin","Transfers","Receipts",
    "Buckets","Issues","Adjustments","Depletions","Operating Inventory","Last Zero Date"
]

MONTHS = {"JAN":"Jan","FEB":"Feb","MAR":"Mar","APR":"Apr","MAY":"May","JUN":"Jun",
          "JUL":"Jul","AUG":"Aug","SEP":"Sep","OCT":"Oct","NOV":"Nov","DEC":"Dec"}

def latest_csv(here: Path) -> Path:
    files = sorted(here.glob('*.csv'), key=lambda p: p.stat().st_mtime, reverse=True)
    if not files:
        raise FileNotFoundError(f"No CSV files found in {here}")
    return files[0]

def to_int(s: str):
    if s is None: return None
    s = s.strip().replace(',', '')
    if s == '': return None
    try:
        return int(s)
    except ValueError:
        return None

def normalize_pile(pile: str):
    if not pile: return None
    m = re.match(r'^[A-Za-z0-9]+', pile.strip())
    return m.group(0) if m else None

def parse_header_date(raw_text: str):
    """Return 'DD-Mon' string from the header line if present, else None."""
    m = re.search(r"created at\s+(\d{1,2})-([A-Za-z]{3})-\d{4}\s+", raw_text, re.IGNORECASE)
    if not m:
        return None
    day = m.group(1)
    mon3 = m.group(2).upper()
    mon = MONTHS.get(mon3, mon3)
    return f"{day}-{mon}"

def parse_csv(path: Path):
    # Read entire file
    raw = path.read_text(encoding='utf-8')
    # Build meta
    meta = {"report_date": parse_header_date(raw)}

    # Use csv.reader for rows
    rows = list(csv.reader(raw.splitlines()))

    # Locate header row by presence of required column
    header_idx = None
    for i, row in enumerate(rows):
        if 'Operating Inventory' in row:
            header_idx = i
            break
    if header_idx is None:
        raise ValueError('Header row not found (missing "Operating Inventory").')
    header = rows[header_idx]

    # Column indices
    idx = {name: header.index(name) for name in EXPECTED if name in header}
    stock = {}
    for row in rows[header_idx+1:]:
        if len(row) < len(header):
            continue
        pile_key = normalize_pile(row[idx['Pile']] if 'Pile' in idx else '')
        if not pile_key:
            continue
        stock[pile_key] = {
            'code': row[idx['Code']] if 'Code' in idx else None,
            'material': row[idx['Material']] if 'Material' in idx else None,
            'pile': pile_key,
            'operating_inventory_lbs': to_int(row[idx['Operating Inventory']] if 'Operating Inventory' in idx else None),
            'month_begin_lbs':         to_int(row[idx['Month Begin']] if 'Month Begin' in idx else None),
            'receipts_lbs':            to_int(row[idx['Receipts']] if 'Receipts' in idx else None),
            'issues_lbs':              to_int(row[idx['Issues']] if 'Issues' in idx else None),
            'transfers_lbs':           to_int(row[idx['Transfers']] if 'Transfers' in idx else None),
            'adjustments_lbs':         to_int(row[idx['Adjustments']] if 'Adjustments' in idx else None),
            'depletions_lbs':          to_int(row[idx['Depletions']] if 'Depletions' in idx else None),
            'last_zero_date':          row[idx['Last Zero Date']] if 'Last Zero Date' in idx else None,
        }
    return stock, meta

def main():
    here = Path.cwd()
    # Determine source CSV
    if len(sys.argv) > 1:
        src = Path(sys.argv[1])
        if not src.exists():
            print(f"ERROR: CSV not found: {src}")
            sys.exit(2)
    else:
        src = latest_csv(here)

    stock, meta = parse_csv(src)
    out_path = here / 'stockData.json'
    payload = {"meta": meta, "stock": stock}
    out_path.write_text(json.dumps(payload, indent=2), encoding='utf-8')
    print(f"Wrote {out_path} with {len(stock)} entries from {src.name}; date={meta.get('report_date')}")

if __name__ == '__main__':
    main()
