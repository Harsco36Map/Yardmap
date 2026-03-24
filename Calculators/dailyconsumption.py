#!/usr/bin/env python3
from pathlib import Path
import sys
import csv
from datetime import datetime

DEFAULT_LBS_PER_TON = 2000

def detect_delimiter(file_path: Path) -> str:
    with file_path.open('r', encoding='utf-8', errors='ignore') as f:
        sample = f.read(4096)
    if not sample.strip():
        return ','
    try:
        dialect = csv.Sniffer().sniff(sample)
        return dialect.delimiter
    except Exception:
        for sep in [',', '\t', ';', '\n']:
            if sep in sample:
                return sep
        return ','

def parse_date(value: str):
    if value is None:
        return None
    v = value.strip()
    if not v:
        return None
    for fmt in ("%Y-%m-%d", "%Y/%m/%d", "%m/%d/%Y", "%m/%d/%y", "%Y.%m.%d"):
        try:
            return datetime.strptime(v, fmt)
        except ValueError:
            pass
    try:
        return datetime.fromisoformat(v)
    except Exception:
        return None

def to_float(value: str):
    if value is None:
        return None
    v = value.strip().replace(',', '')
    if v == '':
        return None
    try:
        return float(v)
    except ValueError:
        return None

def read_rows(file_path: Path, debug: bool = False):
    sep = detect_delimiter(file_path)
    if debug:
        print(f"[debug] Source: {file_path.name!r}, delimiter: {repr(sep)}")
    rows = []
    with file_path.open('r', encoding='utf-8', errors='ignore', newline='') as f:
        reader = csv.reader(f, delimiter=sep)
        for row in reader:
            rows.append(row)
    if debug:
        print(f"[debug] total lines read={len(rows)}")
    return rows

def extract_consumed_date(rows, debug: bool = False):
    out = []
    for row in rows:
        if not row:
            continue
        n = len(row)
        if n < 3:
            continue
        consumed_idx = n - 3  # J
        date_idx = n - 2      # K
        consumed = to_float(row[consumed_idx]) if consumed_idx < n else None
        dt = parse_date(row[date_idx]) if date_idx < n else None
        if consumed is not None and dt is not None:
            out.append((dt, consumed))
    if debug:
        print(f"[debug] valid rows (daily) after type coercion={len(out)}")
    return out

def filter_month(pairs, month, debug: bool = False):
    if not month:
        return pairs
    try:
        start = datetime.strptime(str(month) + "-01", "%Y-%m-%d")
    except ValueError:
        if debug:
            print(f"[warn] Invalid --month {month!r}; skipping filter")
        return pairs
    if start.month == 12:
        nxt = datetime(year=start.year + 1, month=1, day=1)
    else:
        nxt = datetime(year=start.year, month=start.month + 1, day=1)
    end = nxt
    filtered = [(dt, c) for (dt, c) in pairs if start <= dt < end]
    if debug:
        print(f"[debug] month={month}, kept={len(filtered)} of {len(pairs)}")
    return filtered

def group_daily(pairs, lbs_per_ton: int):
    totals = {}
    for dt, c in pairs:
        d = dt.date()
        totals[d] = totals.get(d, 0.0) + c
    out = []
    for d in sorted(totals.keys()):
        consumed = totals[d]
        net_tons = consumed / float(lbs_per_ton)
        out.append([d.isoformat(), f"{consumed:.0f}", f"{net_tons:.6f}"])
    return out

def extract_day_pile_consumed(rows, debug: bool = False):
    out = []
    for row in rows:
        if not row:
            continue
        n = len(row)
        if n < 6:
            continue
        pile_idx = n - 6  # G
        consumed_idx = n - 3  # J
        date_idx = n - 2  # K
        pile_raw = row[pile_idx] if pile_idx < n else None
        cons_raw = row[consumed_idx] if consumed_idx < n else None
        day_raw = row[date_idx] if date_idx < n else None
        dt = parse_date(day_raw)
        pile_label = (str(pile_raw).strip() if pile_raw is not None else None)
        consumed = to_float(cons_raw)
        if dt is not None and pile_label and consumed is not None:
            out.append((dt, pile_label, consumed))
    if debug:
        print(f"[debug] valid rows for pile totals={len(out)} (Pile # from column G=n-6)")
    return out

def filter_month_triples(triples, month, debug: bool = False):
    if not month:
        return triples
    try:
        start = datetime.strptime(str(month) + "-01", "%Y-%m-%d")
    except ValueError:
        if debug:
            print(f"[warn] Invalid --month {month!r}; skipping pile filter")
        return triples
    if start.month == 12:
        nxt = datetime(year=start.year + 1, month=1, day=1)
    else:
        nxt = datetime(year=start.year, month=start.month + 1, day=1)
    end = nxt
    filtered = [(dt, pile, a) for (dt, pile, a) in triples if start <= dt < end]
    if debug:
        print(f"[debug] pile month={month}, kept={len(filtered)} of {len(triples)}")
    return filtered

def group_pile_totals(triples):
    pile_totals = {}
    unique_days = set()
    for dt, pile, a in triples:
        unique_days.add(dt.date())
        pile_totals[pile] = pile_totals.get(pile, 0.0) + a
    num_days = len(unique_days)
    out_rows = []
    for pile in sorted(pile_totals.keys(), key=lambda x: str(x)):
        total = pile_totals.get(pile, 0.0)
        avg = (total / num_days) if num_days > 0 else 0.0
        out_rows.append([str(pile), f"{total:.0f}", f"{avg:.2f}"])
    return out_rows, num_days

def main():
    here = Path.cwd()
    csv_arg = None
    month = None
    lbs_per_ton = DEFAULT_LBS_PER_TON
    debug = False

    argv = sys.argv[1:]
    i = 0
    while i < len(argv):
        a = argv[i]
        if a.startswith('--month'):
            if a == '--month' and i + 1 < len(argv):
                month = argv[i + 1]
                i += 2
            else:
                month = a.split('=')[1]
                i += 1
        elif a.startswith('--lbs-per-ton'):
            if a == '--lbs-per-ton' and i + 1 < len(argv):
                lbs_per_ton = int(argv[i + 1])
                i += 2
            else:
                lbs_per_ton = int(a.split('=')[1])
                i += 1
        elif a == '--debug':
            debug = True
            i += 1
        elif a.startswith('--'):
            i += 1
        else:
            csv_arg = a
            i += 1

    if csv_arg:
        src = (here / csv_arg) if not Path(csv_arg).is_absolute() else Path(csv_arg)
        if not src.exists():
            print(f"ERROR: CSV not found: {src}")
            sys.exit(2)
    else:
        # fallback to most recent
        files = list(here.glob('*.csv')) + list(here.glob('*.CSV'))
        files = sorted(files, key=lambda p: p.stat().st_mtime, reverse=True)
        if not files:
            print("ERROR: No CSV files found")
            sys.exit(2)
        src = files[0]

    rows = read_rows(src, debug=debug)

    # Daily
    pairs = extract_consumed_date(rows, debug=debug)
    pairs = filter_month(pairs, month=month, debug=debug)
    out_rows_daily = group_daily(pairs, lbs_per_ton=lbs_per_ton)

    # Pile
    triples = extract_day_pile_consumed(rows, debug=debug)
    triples = filter_month_triples(triples, month=month, debug=debug)
    pile_rows, num_days_in_filtered = group_pile_totals(triples)

    suffix = f"_{month}" if month else ""
    out_path = here / f"{src.stem}_daily_totals{suffix}.csv"

    # Side-by-side output: A-C daily, D spacer, E-G pile
    max_rows = max(len(out_rows_daily), len(pile_rows))
    with out_path.open('w', encoding='utf-8', newline='') as f:
        writer = csv.writer(f)
        writer.writerow(["Day", "Consumed", "Net_Tons", "", "Pile", "Total_Actual", "Avg_Daily"])
        for i in range(max_rows):
            daily = out_rows_daily[i] if i < len(out_rows_daily) else ["", "", ""]
            pile = pile_rows[i] if i < len(pile_rows) else ["", "", ""]
            writer.writerow(daily + [""] + pile)

    print(f"Wrote {out_path!r} with side-by-side daily ({len(out_rows_daily)}) and pile ({len(pile_rows)}) rows; unique_days_for_avg={num_days_in_filtered}")

if __name__ == '__main__':
    main()
