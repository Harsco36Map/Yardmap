/* ===================================================================
 PAST DUE (pastdue.js)

 Everything related to the "past due" pile mechanics lives here:
   - Which piles are exempt (by marker type or by individual pile code)
   - Building the past-due list (> 6 months since last zero date)
   - Piles flagged "stopped":"yes" in markers.json are ALWAYS included,
     even if their date wouldn't qualify them and even if exempted below
   - The PastDue XLSX export (with consumption averages / days-until-depleted)

 Load order: requires utils.js (extractPileCode, normalizePileCode,
 parseMDY, monthsDiff, isStoppedPileCode, markerIsStopped, parseCsvLine,
 getCurrentInventoryPeriod). Must load before map.js.
=================================================================== */

/* -------------------------------------------------------------------
 EXEMPTIONS
------------------------------------------------------------------- */

// Marker types exempt from the age-based past-due mechanics.
// (Alloys are exempt EXCEPT molybdenum oxide — see isPastDueExempt.)
const PASTDUE_EXEMPT_TYPES = new Set(["Coils", "Breaking", "Unbreakable", "Alloys"]);

// Individual pile exemptions — add pile codes here (quoted, comma-separated)
// to exclude specific piles from the age-based past-due calculations.
// Leading zeros don't matter ("092" matches "92"). Example:
//   const PASTDUE_EXEMPT_PILES = ["148", "20H"];
const PASTDUE_EXEMPT_PILES = ["931"
];

const pastDueExemptPileSet = new Set(
  PASTDUE_EXEMPT_PILES.map(c => normalizePileCode(c)).filter(Boolean)
);

function isPastDueExempt(marker, stockIndex) {
  if (!marker) return true;
  const code = extractPileCode(marker.name);
  if (code && pastDueExemptPileSet.has(normalizePileCode(code))) return true;
  if (!PASTDUE_EXEMPT_TYPES.has(marker.type)) return false;
  if (marker.type === "Alloys") {
    const name = String(marker.name || "").toLowerCase();
    const s = code ? stockIndex[code] : null;
    const material = String((s && s.material) || "").toLowerCase();
    const isMoOx =
      name.includes("molybdenum oxide") ||
      material.includes("molybdenum oxide");
    return !isMoOx;
  }
  return true;
}

/* -------------------------------------------------------------------
 AGE FORMATTING
------------------------------------------------------------------- */

function formatAgeYM(totalMonths) {
  if (typeof totalMonths !== 'number' || !isFinite(totalMonths) || totalMonths < 0) return '';
  const years = Math.floor(totalMonths / 12);
  const months = totalMonths % 12;
  const yPart = years > 0 ? `${years} ${years === 1 ? 'year' : 'years'}` : '';
  const mPart = months > 0 ? `${months} ${months === 1 ? 'month' : 'months'}` : (years === 0 ? '0 months' : '');
  return yPart && mPart ? `${yPart} ${mPart}` : (yPart || mPart);
}

/* -------------------------------------------------------------------
 PAST DUE LIST BUILDER
 A pile is included when either:
   - its last zero date is more than 6 months old and it isn't exempt, or
   - it is flagged "stopped":"yes" in markers.json (always included;
     overrides both type and individual exemptions).
 Only one entry per pile code (main/remote virtual piles share one row).
------------------------------------------------------------------- */

function buildPastDueList(markers, stockIndex) {
  const pastDue = [];
  const seenCodes = new Set();

  (Array.isArray(markers) ? markers : []).forEach(marker => {
    const code = extractPileCode(marker.name);
    if (!code || seenCodes.has(code)) return;

    const s = stockIndex[code];
    const dt = s ? parseMDY(s.last_zero_date) : null;
    const mAge = dt ? monthsDiff(dt, new Date()) : null;

    const agePastDue = mAge !== null && mAge > 6 && !isPastDueExempt(marker, stockIndex);
    const stopped = isStoppedPileCode(code) || markerIsStopped(marker);
    if (!agePastDue && !stopped) return;

    pastDue.push({
      code: code,
      name: marker.name,
      rawType: marker.type,
      material: (s && s.material) || '',
      lastZero: (s && s.last_zero_date) || '',
      ageMonths: mAge,
      ageLabel: mAge !== null ? formatAgeYM(mAge) : '',
      invLbs: s ? s.operating_inventory_lbs : null,
      stopped: stopped,
      marker: marker._leaflet
    });
    seenCodes.add(code);
  });

  pastDue.sort((a, b) =>
    ((b.ageMonths ?? -1) - (a.ageMonths ?? -1)) || a.code.localeCompare(b.code)
  );
  window.pastDue = pastDue;
  return pastDue;
}

/* -------------------------------------------------------------------
 CONSUMPTION CSV PARSER (monthly averages for the export math)
------------------------------------------------------------------- */

const consumptionCsvUrl = 'AverageConsumption.csv'; // <-- set per month

async function fetchConsumptionCsv() {
  const res = await fetch(consumptionCsvUrl);
  if (!res.ok) throw new Error(`Failed to fetch ${consumptionCsvUrl}: ${res.status}`);
  const text = await res.text();
  // New fixed format:
  // A: Pile Number | B: Material Lot # | C: Total Lbs consumed | D: Total Tons consumed
  // E1: Days represented by this monthly total (manually maintained)
  const lines = text.split(/\r?\n/).filter(l => String(l || '').trim() !== '');
  if (!lines.length) throw new Error("Consumption CSV is empty.");

  const headerParts = parseCsvLine(lines[0]).map(v => String(v || '').trim());
  const headerNorm = headerParts.map(v => v.toLowerCase());
  const findCol = (...names) => {
    const wanted = names.map(n => String(n).toLowerCase());
    return headerNorm.findIndex(h => wanted.includes(h));
  };

  const pileCol = findCol('pile_number', 'pile number', 'pile', 'pile #', 'pile#');
  const totalLbsCol = findCol('total_lbs', 'total lbs', 'total_lbs consumed', 'total lbs consumed');
  const resolvedPileCol = pileCol >= 0 ? pileCol : 0;
  const resolvedTotalLbsCol = totalLbsCol >= 0 ? totalLbsCol : 2;

  const { year, month } = getCurrentInventoryPeriod();
  const daysInMonth = new Date(year, month + 1, 0).getDate();
  const configuredDays = Number(String(headerParts[4] || '').trim());
  const dayDivisor = (Number.isFinite(configuredDays) && configuredDays > 0)
    ? configuredDays
    : Math.max(1, daysInMonth);

  const pileTotalsByCode = {}; // { CODE -> totalLbs }
  for (let i = 1; i < lines.length; i++) {
    const raw = lines[i];
    if (!raw) continue;
    const parts = parseCsvLine(raw);
    if (parts.length < 3) continue;

    const pile = normalizePileCode(parts[resolvedPileCol]);
    if (!pile) continue;

    const totalLbs = Number(String(parts[resolvedTotalLbsCol] || '').replace(/,/g, '').trim());
    if (!Number.isFinite(totalLbs)) continue;

    pileTotalsByCode[pile] = (pileTotalsByCode[pile] || 0) + totalLbs;
  }

  const pileAvgByCode = {};
  Object.entries(pileTotalsByCode).forEach(([code, totalLbs]) => {
    pileAvgByCode[code] = totalLbs / dayDivisor;
  });

  return { pileAvgByCode };
}

window.fetchConsumptionCsv = fetchConsumptionCsv;

/* -------------------------------------------------------------------
 PAST DUE XLSX EXPORT (robust, non-corrupt)
------------------------------------------------------------------- */

window.exportPastDueXlsx = async function exportPastDueXlsx() {
  try {
    const XLSXlib = window.XLSX;
    if (!XLSXlib || !XLSXlib.utils) {
      alert('XLSX library not loaded.');
      console.error('XLSX missing or invalid:', window.XLSX);
      return;
    }

    // Fetch consumption averages
    const { pileAvgByCode } = await fetchConsumptionCsv();

    const rows = (window.pastDue || []).map(p => {
      const codeKey = normalizePileCode(p.code);
      const invLbs = typeof p.invLbs === 'number' ? Math.max(0, p.invLbs) : 0;
      const avgDaily = pileAvgByCode[codeKey] ?? 0;
      const dud = avgDaily > 0 ? invLbs / avgDaily : 0;

      return {
        'Pile Number': p.code ?? '—',
        'Name': p.name ?? '—',
        'Material': p.material ?? '—',
        'Last Zero Date': p.lastZero || '—',
        'Age': p.ageLabel ?? '—',
        'Stopped': p.stopped ? 'Yes' : '',
        'Inventory (lbs)': invLbs,
        'Average Consumed Daily': Math.round(avgDaily || 0),
        'Days Until Depleted': Number.isFinite(dud) ? Number(dud.toFixed(1)) : 0,
        'Action': ''
      };
    });

    // Convert to worksheet
    const ws = XLSXlib.utils.json_to_sheet(rows);

    // Auto column widths
    const colWidths = Object.keys(rows[0] || {}).map(k => ({
      wch: Math.max(
        k.length,
        ...rows.map(r => String(r[k] ?? '').length)
      ) + 2
    }));
    ws['!cols'] = colWidths;

    // Build workbook
    const wb = XLSXlib.utils.book_new();
    XLSXlib.utils.book_append_sheet(wb, ws, 'PastDue');

    // Filename
    const today = new Date();
    const yyyy = today.getFullYear();
    const mm = String(today.getMonth() + 1).padStart(2, '0');
    const dd = String(today.getDate()).padStart(2, '0');

    XLSXlib.writeFile(wb, `PastDue_${yyyy}-${mm}-${dd}.xlsx`);

  } catch (err) {
    console.error('Past Due export failed:', err);
    alert('Past Due export failed.');
  }
};
