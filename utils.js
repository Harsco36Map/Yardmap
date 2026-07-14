/* ===================================================================
   utils.js  —  Shared helpers for Yardmap
   Loaded before all station files and map.js via <script> tags.
   All functions go into global scope (no import/export needed).
   =================================================================== */

// ─── Number / format helpers ──────────────────────────────────────────
function fmtInt(n)  { return (typeof n === 'number' && isFinite(n)) ? Math.round(n).toLocaleString('en-US') : '—'; }
function parseWeight(s) {
  if (!s) return NaN;
  const n = Number(String(s).replace(/,/g, '').trim());
  return Number.isFinite(n) ? n : NaN;
}
function fmtTons(n, d = 3) { return (typeof n === 'number' && isFinite(n)) ? n.toFixed(d) : '—'; }
function fmtTons2(n, d = 2) { return (typeof n === 'number' && isFinite(n)) ? n.toFixed(d) : '—'; }

// ─── Date helpers ─────────────────────────────────────────────────────
function parseInventoryReportDate(dateStr) {
  if (!dateStr) return null;
  let parsed = null;
  const s = String(dateStr).trim();

  // DD-MMM-YYYY or DD-MMM-YYYY HH:MM:SS
  const dmy = s.match(/^(\d{1,2})-([A-Za-z]{3})-(\d{4})/);
  if (dmy) {
    const day = Number(dmy[1]);
    const m = dmy[2].toUpperCase();
    const year = Number(dmy[3]);
    const monthMap = {
      JAN: 0, FEB: 1, MAR: 2, APR: 3, MAY: 4, JUN: 5,
      JUL: 6, AUG: 7, SEP: 8, OCT: 9, NOV: 10, DEC: 11
    };
    if (!Number.isNaN(day) && !Number.isNaN(year) && monthMap.hasOwnProperty(m)) {
      parsed = new Date(year, monthMap[m], day);
    }
  }

  if (!parsed) {
    const mdy = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})/);
    if (mdy) {
      const month = Number(mdy[1]) - 1;
      const day = Number(mdy[2]);
      const year = Number(mdy[3]);
      if (!Number.isNaN(month) && !Number.isNaN(day) && !Number.isNaN(year)) {
        parsed = new Date(year, month, day);
      }
    }
  }

  if (!parsed) {
    const dt = new Date(s);
    if (isFinite(dt.getTime())) { parsed = dt; }
  }

  return parsed;
}

function getCurrentInventoryPeriod() {
  const year = Number.isInteger(latestInventoryPeriod.year) ? latestInventoryPeriod.year : new Date().getFullYear();
  const month = Number.isInteger(latestInventoryPeriod.month) ? latestInventoryPeriod.month : new Date().getMonth();
  return { year, month };
}

function getCurrentInventoryMonthLabel() {
  const { year, month } = getCurrentInventoryPeriod();
  return new Date(year, month, 1).toLocaleString('en-US', { month: 'long' });
}

function formatInventoryBannerDate(dateStr) {
  const parsed = parseInventoryReportDate(dateStr);
  if (!parsed) return String(dateStr || '—').trim() || '—';
  return `${parsed.toLocaleString('en-US', { month: 'long' })}-${parsed.getDate()}`;
}

function getDateWeekdayLetter(dateValue) {
  let parsed = null;
  if (dateValue instanceof Date && isFinite(dateValue.getTime())) {
    parsed = dateValue;
  } else if (typeof dateValue === 'string' && dateValue.trim()) {
    parsed = parseXlsxDateCell(dateValue);
  }

  if (!(parsed instanceof Date) || !isFinite(parsed.getTime())) return '';

  const weekdayLetters = ['S', 'M', 'T', 'W', 'T', 'F', 'S'];
  return weekdayLetters[parsed.getDay()] || '';
}

function getDateWeekdayColor(dateValue) {
  let parsed = null;
  if (dateValue instanceof Date && isFinite(dateValue.getTime())) {
    parsed = dateValue;
  } else if (typeof dateValue === 'string' && dateValue.trim()) {
    parsed = parseXlsxDateCell(dateValue);
  }

  if (!(parsed instanceof Date) || !isFinite(parsed.getTime())) return '#666';

  const day = parsed.getDay();
  if (day === 0) return '#c62828';
  if (day === 6) return '#ef6c00';
  if (day === 3) return '#2e7d32';
  return '#0b57d0';
}

function renderDateWithWeekday(dateValue, label, options = {}) {
  const text = esc(label || '—');
  const letter = getDateWeekdayLetter(dateValue);
  if (!letter) return text;

  const {
    button = false,
    buttonClass = '',
    dataDate = '',
    extraButtonStyle = ''
  } = options;

  const badge = `<span style="display:inline-block;min-width:14px;margin-right:6px;font-weight:700;color:${getDateWeekdayColor(dateValue)}">${letter}</span>`;
  const content = `${badge}<span>${text}</span>`;

  if (!button) return content;

  return `<button type="button" class="${buttonClass}" data-date="${esc(dataDate)}" style="border:none;background:none;color:#0b57d0;text-decoration:underline;padding:0;cursor:pointer;font:inherit;display:inline-flex;align-items:center;${extraButtonStyle}">${content}</button>`;
}

function parseXlsxDateCell(value) {
  if (value instanceof Date && isFinite(value.getTime())) return value;
  const v = String(value || '').trim();
  if (!v) return null;

  // Parse YYYY-MM-DD as a local date (avoid UTC shift from new Date('YYYY-MM-DD')).
  const isoDateOnly = v.match(/^(\d{4})-(\d{1,2})-(\d{1,2})$/);
  if (isoDateOnly) {
    const year = Number(isoDateOnly[1]);
    const month = Number(isoDateOnly[2]) - 1;
    const day = Number(isoDateOnly[3]);
    if (!Number.isNaN(year) && !Number.isNaN(month) && !Number.isNaN(day)) {
      return new Date(year, month, day);
    }
  }

  // Prefer inventory parser first
  const parsed = parseInventoryReportDate(v);
  if (parsed) return parsed;
  const dt = new Date(v);
  return isFinite(dt.getTime()) ? dt : null;
}

function findRepeatedBlockStartIndexes(headers, template) {
  const canon = headers.map(h => String(h || '').trim().toLowerCase());
  const templateL = template.map(t => String(t || '').trim().toLowerCase());
  const starts = [];
  for (let c = 0; c <= canon.length - templateL.length; c++) {
    let matched = true;
    for (let j = 0; j < templateL.length; j++) {
      if (canon[c + j] !== templateL[j]) {
        matched = false;
        break;
      }
    }
    if (matched) starts.push(c);
  }
  return starts;
}

function downloadCsvFromRows(filename, rows, headers) {
  if (!Array.isArray(rows)) return;
  const escape = val => {
    const text = String(val ?? '');
    return '"' + text.replace(/"/g, '""') + '"';
  };

  const lines = [headers.map(escape).join(',')];
  for (const r of rows) {
    const row = headers.map(h => escape(r[h]));
    lines.push(row.join(','));
  }

  const blob = new Blob([lines.join('\r\n')], { type: 'text/csv;charset=utf-8;' });
  const a = document.createElement('a');
  a.href = URL.createObjectURL(blob);
  a.download = filename;
  a.style.display = 'none';
  document.body.appendChild(a);
  a.click();
  URL.revokeObjectURL(a.href);
  document.body.removeChild(a);
}

// ─── XLSX / workbook helpers ──────────────────────────────────────────
function findSheetByName(wb, name) {
  if (!wb || !Array.isArray(wb.SheetNames)) return null;
  const found = wb.SheetNames.find(n => n.trim().toLowerCase() === String(name).trim().toLowerCase());
  return found ? wb.Sheets[found] : null;
}

async function loadTotalsWorkbook(force = false) {
  const now = Date.now();
  if (!force && totalsWorkbookCache.workbook && (now - totalsWorkbookCache.at) < 120000) {
    return totalsWorkbookCache.workbook;
  }

  if (!window.XLSX || !window.XLSX.read) {
    throw new Error('XLSX library not available');
  }

  const res = await fetch(totalsXlsxUrl, { cache: 'no-store' });
  if (!res.ok) throw new Error(`Failed to fetch ${totalsXlsxUrl}: ${res.status}`);
  const data = await res.arrayBuffer();
  const workbook = window.XLSX.read(data, { type: 'array' });

  totalsWorkbookCache = { at: now, workbook };
  return workbook;
}

async function loadHistoryWorkbook(force = false) {
  const now = Date.now();
  if (!force && historyWorkbookCache.workbook && (now - historyWorkbookCache.at) < 120000) {
    return historyWorkbookCache.workbook;
  }
  if (!window.XLSX || !window.XLSX.read) throw new Error('XLSX library not available');
  const res = await fetch('History.xlsx', { cache: 'no-store' });
  if (!res.ok) throw new Error(`Failed to fetch History.xlsx: ${res.status}`);
  const buf = await res.arrayBuffer();
  const workbook = window.XLSX.read(buf, { type: 'array' });
  historyWorkbookCache = { at: now, workbook };
  return workbook;
}

function formatMonthYear(year, month) {
  const names = ['January','February','March','April','May','June','July','August','September','October','November','December'];
  return (names[month] || 'Month') + ' ' + year;
}

async function discoverHistoryMonths(baseSheetType) {
  // History.xlsx uses sheet names like "1-2026Receiving1" or "3-2026Consumption1".
  // Discover available months by matching that pattern instead of scanning row data.
  try {
    const workbook = await loadHistoryWorkbook();
    if (!workbook || !Array.isArray(workbook.SheetNames)) return [];
    const pattern = new RegExp(`^(\\d{1,2})-(\\d{4})${baseSheetType}$`, 'i');
    const seen = new Map();
    for (const name of workbook.SheetNames) {
      const m = name.trim().match(pattern);
      if (m) {
        const month = parseInt(m[1], 10) - 1; // convert to 0-indexed
        const year = parseInt(m[2], 10);
        const key = `${year}-${month}`;
        if (!seen.has(key)) seen.set(key, { year, month });
      }
    }
    return Array.from(seen.values()).sort((a, b) => b.year - a.year || b.month - a.month);
  } catch (err) {
    console.warn('discoverHistoryMonths failed:', err);
    return [];
  }
}

// Discovers months for panels whose history lives in year-named sheets (e.g. "2026BurningHistory").
// Each sheet stores months as horizontal blocks that each have their own "Date" header column.
// We find those header columns and sample one date per block — avoids false positives from
// weight/quantity values in other columns.
async function discoverHistoryMonthsForYearlySheet(basePattern) {
  try {
    const workbook = await loadHistoryWorkbook();
    if (!workbook || !Array.isArray(workbook.SheetNames)) return [];
    const pat = new RegExp(`^\\d{4}${basePattern}$`, 'i');
    const seen = new Map();
    for (const shName of workbook.SheetNames) {
      if (!pat.test(shName.trim())) continue;
      const sheet = workbook.Sheets[shName];
      if (!sheet) continue;
      const data = window.XLSX.utils.sheet_to_json(sheet, { header: 1, raw: false });
      if (!Array.isArray(data) || data.length < 2) continue;

      // Find every column whose header is exactly "date" — one per horizontal month block.
      const headers = (Array.isArray(data[0]) ? data[0] : []).map(h => String(h || '').trim().toLowerCase());
      const dateCols = [];
      headers.forEach((h, i) => { if (h === 'date') dateCols.push(i); });
      if (dateCols.length === 0) dateCols.push(0); // fallback: first column

      // For each date column, take the first valid date found — that identifies the month.
      for (const col of dateCols) {
        for (let r = 1; r < data.length; r++) {
          const row = Array.isArray(data[r]) ? data[r] : [];
          const parsed = parseXlsxDateCell(row[col]);
          if (parsed && !isNaN(parsed.getTime()) && parsed.getFullYear() >= 2020) {
            const key = `${parsed.getFullYear()}-${parsed.getMonth()}`;
            if (!seen.has(key)) seen.set(key, { year: parsed.getFullYear(), month: parsed.getMonth() });
            break; // one date per block is enough
          }
        }
      }
    }
    return Array.from(seen.values()).sort((a, b) => b.year - a.year || b.month - a.month);
  } catch (err) {
    console.warn('discoverHistoryMonthsForYearlySheet failed:', err);
    return [];
  }
}

// Returns the correct sheet name for a given base type and optional period override.
// Production.xlsx uses bare names ("Consumption1"); History.xlsx uses "1-2026Consumption1".
function historySheetName(baseType, periodOverride) {
  if (!periodOverride) return baseType;
  return `${periodOverride.month + 1}-${periodOverride.year}${baseType}`;
}

// ─── Pile / marker helpers ────────────────────────────────────────────
// Leading alphanumerics (e.g., "62U" from "62U Unbreakable")
function extractPileCode(name) {
  if (!name) return null;
  const m = String(name).match(/^[A-Za-z0-9]+/);
  return m ? m[0] : null;
}

// Normalize pile codes so numeric codes match regardless of leading zeros (e.g., 092 === 92).
function normalizePileCode(code) {
  const raw = String(code ?? '').trim().toUpperCase();
  if (!raw) return '';
  if (/^\d+$/.test(raw)) {
    return String(parseInt(raw, 10));
  }
  return raw;
}

function markerIsStopped(marker) {
  if (!marker || typeof marker !== 'object') return false;
  const raw = marker.stopped ?? marker.Stopped ?? marker.STOPPED;
  if (raw === true) return true;
  if (typeof raw === 'number') return raw !== 0;
  const text = String(raw ?? '').trim().toLowerCase();
  return text === 'yes' || text === 'y' || text === 'true' || text === '1' || text === 'stopped';
}

function rebuildStoppedPileCodes(markers) {
  const out = new Set();
  (Array.isArray(markers) ? markers : []).forEach(marker => {
    if (!markerIsStopped(marker)) return;
    const code = normalizePileCode(extractPileCode(marker.name));
    if (code) out.add(code);
  });
  stoppedPileCodesGlobal = out;
}

function isStoppedPileCode(code) {
  const key = normalizePileCode(code);
  return !!key && stoppedPileCodesGlobal.has(key);
}

function normalizeLookupKey(value) {
  return String(value ?? '').trim().toUpperCase();
}

const receivingPileMaterialOverrides = {
  '961': 'Dolomitic Lime'
};

function getMaterialFromMarkerNameByPile(pileKey) {
  if (!pileKey) return '';
  const marker = (allMarkersData || []).find(m => {
    const markerCode = normalizePileCode(extractPileCode(m && m.name));
    return markerCode && markerCode === pileKey;
  });

  if (!marker || !marker.name) return '';
  const markerName = String(marker.name).trim();
  const code = extractPileCode(markerName);
  if (!code) return markerName;

  const cleaned = markerName.replace(new RegExp(`^${code}\\s*[-:]*\\s*`, 'i'), '').trim();
  return cleaned || markerName;
}

function getMaterialForLotOrPile(lot, pile) {
  const lotKey = normalizeLookupKey(lot);
  const pileKey = normalizePileCode(pile);

  if (lotKey) {
    const byLot = Object.values(stockIndexGlobal || {}).find(item =>
      normalizeLookupKey(item && item.code) === lotKey
    );
    if (byLot && byLot.material) return String(byLot.material);
  }

  if (pileKey && stockIndexGlobal && stockIndexGlobal[pileKey] && stockIndexGlobal[pileKey].material) {
    return String(stockIndexGlobal[pileKey].material);
  }

  if (pileKey) {
    const fromMarkerName = getMaterialFromMarkerNameByPile(pileKey);
    if (fromMarkerName) return fromMarkerName;
  }

  if (pileKey && receivingPileMaterialOverrides[pileKey]) {
    return receivingPileMaterialOverrides[pileKey];
  }

  return '';
}

// ─── Charge bucket helpers (Consumption4 heat viewers) ────────────────
// Distinct color per charge bucket within a heat, readable on light and dark
// backgrounds as a filled badge with white text.
const BUCKET_SEQ_COLORS = ['#3b82f6', '#f59e0b', '#10b981', '#ef4444', '#a855f7', '#06b6d4', '#f97316', '#84cc16'];

function bucketColorForSeq(seq) {
  if (!Number.isFinite(seq) || seq < 1) return '';
  return BUCKET_SEQ_COLORS[(seq - 1) % BUCKET_SEQ_COLORS.length];
}

// Small colored badge identifying which charge bucket a material entry was
// loaded into. Old-format rows have no bucket info → muted dash.
function bucketBadgeHtml(mat) {
  const seq = mat && mat.bucketSeq;
  if (!Number.isFinite(seq)) return '<span style="color:#94a3b8">—</span>';
  const color = bucketColorForSeq(seq);
  const title = mat.bucketNumber ? `Charge bucket ${seq} · Bucket #${mat.bucketNumber}` : `Charge bucket ${seq}`;
  return `<span class="bkt-badge" style="background:${color}" title="${esc(title)}">B${seq}</span>`;
}

// ─── Inventory status helpers ─────────────────────────────────────────
// MM/DD/YYYY → Date
function parseMDY(dateStr) {
  if (!dateStr) return null;
  const parts = String(dateStr).split('/');
  if (parts.length !== 3) return null;
  const m = parseInt(parts[0], 10);
  const d = parseInt(parts[1], 10);
  const y = parseInt(parts[2], 10);
  if (!m || !d || !y) return null;
  return new Date(y, m - 1, d);
}

// Full months difference (a → b)
function monthsDiff(a, b) {
  const years = b.getFullYear() - a.getFullYear();
  const months = years * 12 + (b.getMonth() - a.getMonth());
  return (b.getDate() >= a.getDate()) ? months : months - 1;
}

// Last-zero color policy
function lastZeroColor(dateStr) {
  const dt = parseMDY(dateStr);
  if (!dt) return '';
  const now = new Date();
  const m = monthsDiff(dt, now);
  if (m > 6) return 'color:#c62828; font-weight:700;'; // red + bold
  if (m > 3) return 'color:#f9a825; font-weight:700;'; // yellow + bold
  return '';
}

// Bold red when inventory <= 0  (used by Breaking & Burning)
function invAlertStyle(weightLbs) {
  const v = (typeof weightLbs === 'number') ? weightLbs : Number(weightLbs || 0);
  return (Number.isFinite(v) && v <= 0) ? 'color:#c62828;font-weight:700;' : '';
}

function esc(s) {
  return String(s ?? '')
    .replace(/&/g,'&amp;')
    .replace(/</g,'&lt;')
    .replace(/>/g,'&gt;');
}

// IMPORTANT: correct decode
function unescapeAngles(s) {
  return String(s)
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>');
}

// ─── CSV helpers ──────────────────────────────────────────────────────
function parseCsvLine(line) {
  const out = [];
  let cur = '';
  let inQuotes = false;

  for (let i = 0; i < line.length; i++) {
    const ch = line[i];

    if (ch === '"') {
      inQuotes = !inQuotes;
      continue;
    }

    if (ch === ',' && !inQuotes) {
      out.push(cur);
      cur = '';
    } else {
      cur += ch;
    }
  }
  out.push(cur);
  return out.map(v => v.trim());
}

async function fetchLatestInventoryCsv() {
  const res = await fetch('LatestInventory.csv', { cache: 'no-store' });
  if (!res.ok) throw new Error('Failed to load LatestInventory.csv');

  const text = await res.text();
  const lines = text.split(/\r?\n/).filter(Boolean);

  // Extract report date from header
  let report_date = '—';
  for (const line of lines) {
    const m = line.match(/created at\s+(.*)$/i);
    if (m) {
      report_date = m[1].trim();
      break;
    }
  }

  // Find CSV header row
  const headerIndex = lines.findIndex(l => l.startsWith(',"Code"'));
  if (headerIndex === -1) {
    throw new Error('Inventory CSV header not found');
  }

  const stock = {};
  let totalInventoryLbs = 0;

  const num = v => {
    const x = Number((v ?? '').replace(/,/g, '').trim());
    return Number.isFinite(x) ? x : 0;
  };

  for (let i = headerIndex + 1; i < lines.length; i++) {
    const raw = lines[i];
    if (!raw) continue;

    const p = parseCsvLine(raw);
    if (p.length < 12) continue;

    const [
      , code, material, pile,
      monthBegin, transfers, receipts, buckets,
      issues, adjustments, depletions,
      operatingInventory, lastZero
    ] = p;

    if (!pile) continue;

    stock[pile] = {
      code,
      material,
      pile,
      month_begin_lbs: num(monthBegin),
      transfers_lbs: num(transfers),
      receipts_lbs: num(receipts),
      issues_lbs: num(issues),
      adjustments_lbs: num(adjustments),
      depletions_lbs: num(depletions),
      operating_inventory_lbs: num(operatingInventory),
      last_zero_date: lastZero || ''
    };

    totalInventoryLbs += stock[pile].operating_inventory_lbs;
  }

  const parsedInventoryDate = parseInventoryReportDate(report_date);
  if (parsedInventoryDate) {
    latestInventoryPeriod.year = parsedInventoryDate.getFullYear();
    latestInventoryPeriod.month = parsedInventoryDate.getMonth();
  } else {
    const now = new Date();
    latestInventoryPeriod.year = now.getFullYear();
    latestInventoryPeriod.month = now.getMonth();
  }

  return {
    meta: { report_date, total_inventory_lbs: totalInventoryLbs },
    stock
  };
}

// ─── Month selector UI (shared by all five station popups) ────────────
function buildMonthSelectorHtml(prefix, currentPeriod, historyMonths) {
  const { year: curYear, month: curMonth } = getCurrentInventoryPeriod();
  const currentLabel = getCurrentInventoryMonthLabel();
  const displayLabel = currentPeriod ? formatMonthYear(currentPeriod.year, currentPeriod.month) : currentLabel;

  if (historyMonths === null) {
    return `<button type="button" id="${prefix}MonthBtn" style="font-size:11px;font-weight:700;letter-spacing:0.08em;text-transform:uppercase;color:#666;background:none;border:none;padding:0;cursor:pointer;text-decoration:underline dotted;text-underline-offset:2px">${displayLabel} &#x25BE;</button>`;
  }

  const MONTH_SHORT = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];
  const yearSet = new Set([curYear]);
  historyMonths.forEach(m => yearSet.add(m.year));
  const years = [...yearSet].sort((a, b) => a - b);

  const gridsHtml = years.map(year => {
    const cells = MONTH_SHORT.map((abbr, mIdx) => {
      const isCurrentAndSelected = !currentPeriod && year === curYear && mIdx === curMonth;
      const isHistorySelected = currentPeriod && currentPeriod.year === year && currentPeriod.month === mIdx;
      const isCurrentMonth = year === curYear && mIdx === curMonth;
      const hasHistory = historyMonths.some(m => m.year === year && m.month === mIdx);
      const isAvailable = isCurrentMonth || hasHistory;

      if (!isAvailable) {
        return `<span style="display:block;padding:3px 0;font-size:11px;color:#ccc;text-align:center;border-radius:3px">${abbr}</span>`;
      }

      let bg = 'transparent', fw = '400', border = '1px solid transparent', color = '#0b57d0';
      if (isCurrentAndSelected)  { bg = '#e6f4ea'; border = '1px solid #34a853'; color = '#1a6630'; fw = '700'; }
      else if (isHistorySelected) { bg = '#e8f0fe'; border = '1px solid #4285f4'; fw = '700'; }
      else if (isCurrentMonth)    { bg = '#f0fdf4'; }

      const val = isCurrentMonth ? 'current' : `${year}-${mIdx}`;
      return `<button type="button" class="${prefix}MonthCell" data-value="${val}" style="display:block;width:100%;padding:3px 0;font-size:11px;font-weight:${fw};color:${color};background:${bg};border:${border};border-radius:3px;cursor:pointer;text-align:center">${abbr}</button>`;
    });

    return `<div style="text-align:center;font-size:10px;font-weight:700;color:#888;letter-spacing:0.05em;margin-bottom:3px">${year}</div>
      <div style="display:grid;grid-template-columns:repeat(4,1fr);gap:2px;margin-bottom:2px">${cells.join('')}</div>`;
  }).join('');

  const resetBtn = currentPeriod
    ? `<button type="button" id="${prefix}MonthResetBtn" title="Return to current month" style="font-size:15px;line-height:1;background:none;border:none;padding:0 0 0 5px;cursor:pointer;color:#1a73e8;vertical-align:middle">↺</button>`
    : '';

  return `<div style="position:relative;display:inline-block">
      <div style="display:inline-flex;align-items:center;gap:2px">
        <button type="button" id="${prefix}MonthCalendarToggle" style="font-size:11px;font-weight:700;letter-spacing:0.08em;text-transform:uppercase;color:#666;background:none;border:none;padding:0;cursor:pointer;text-decoration:underline dotted;text-underline-offset:2px">${displayLabel}</button>${resetBtn}
      </div>
      <div id="${prefix}MonthCalendar" style="display:none;position:absolute;top:calc(100% + 3px);left:0;z-index:10000;padding:6px 8px;background:#fff;border:1px solid #ddd;border-radius:4px;box-shadow:0 3px 12px rgba(0,0,0,0.2);min-width:160px">
        ${gridsHtml}
      </div>
    </div>`;
}
