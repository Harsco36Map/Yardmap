console.log('XLSX at map.js load:', window.XLSX);

/* ===================================================================
 MAP SETUP
=================================================================== */
var map = L.map('map', {
    center: [40.793903, -82.536906],
    zoom: 18,
    minZoom: 18,
    maxZoom: 21,
    wheelPxPerZoomLevel: 100,
    zoomSnap: 0,
    zoomDelta: 0.25,
    zoomControl: false
});
map.doubleClickZoom.disable();

L.imageOverlay('Scrapyard.png', [
  [40.79156379934851, -82.54114438096362],
  [40.79571031220616, -82.5323681932691]
]).addTo(map);

/* ===================================================================
 Monthly consumption CSV
=================================================================== */
const consumptionCsvUrl = 'averageconsumption.csv'; // <-- set per month

/* ===================================================================
 ICON DEFINITIONS
=================================================================== */
const iconst181 = L.icon({ iconUrl:'icons/st181.png',iconSize:[96,96],iconAnchor:[48,48] });
const iconst430 = L.icon({ iconUrl:'icons/st430.png',iconSize:[96,96],iconAnchor:[48,48] });
const iconst435 = L.icon({ iconUrl:'icons/st435.png',iconSize:[96,96],iconAnchor:[48,48] });
const iconst436 = L.icon({ iconUrl:'icons/st436.png',iconSize:[96,96],iconAnchor:[48,48] });
const iconnickel409 = L.icon({ iconUrl:'icons/nickel409.png',iconSize:[96,96],iconAnchor:[48,48] });
const iconcc409 = L.icon({ iconUrl:'icons/cc409.png',iconSize:[96,96],iconAnchor:[48,48] });
const iconst409 = L.icon({ iconUrl:'icons/st409.png',iconSize:[96,96],iconAnchor:[48,48] });
const iconblend181 = L.icon({ iconUrl:'icons/blend181.png',iconSize:[80,80],iconAnchor:[32,32] });
const iconblend430 = L.icon({ iconUrl:'icons/blend430.png',iconSize:[96,96],iconAnchor:[48,48] });
const iconHome = L.icon({ iconUrl:'icons/Home.png',iconSize:[96,96],iconAnchor:[48,48] });
const iconTundish = L.icon({ iconUrl:'icons/Tundish.png',iconSize:[128,128],iconAnchor:[64,64] });
const iconReclaim = L.icon({ iconUrl:'icons/Reclaim.png',iconSize:[96,96],iconAnchor:[48,48] });
const iconMtown = L.icon({ iconUrl:'icons/Mtown.png',iconSize:[96,96],iconAnchor:[48,48] });
const iconHBI = L.icon({ iconUrl:'icons/HBI.png',iconSize:[128,128],iconAnchor:[64,64] });
const iconFrag = L.icon({ iconUrl:'icons/Frag.png',iconSize:[80,80],iconAnchor:[32,32] });
const iconAlloys = L.icon({ iconUrl:'icons/Alloys.png',iconSize:[32,32],iconAnchor:[16,16] });
const iconOther = L.icon({ iconUrl:'icons/Other.png',iconSize:[96,96],iconAnchor:[48,48] });

/* ===================================================================
 MATERIAL MARKER CONFIG
=================================================================== */
const markerConfig = {
  "st181": { icon: iconst181, displayName: "181 Stainless" },
  "st430": { icon: iconst430, displayName: "430 Stainless" },
  "st435": { icon: iconst435, displayName: "435 Stainless" },
  "st436": { icon: iconst436, displayName: "436 Stainless" },
  "nickel409": { icon: iconnickel409, displayName: "409 Nickel" },
  "cc409": { icon: iconcc409, displayName: "409 Converters" },
  "st409": { icon: iconst409, displayName: "409 Scrap" },
  "blend181": { icon: iconblend181, displayName: "181 Blend" },
  "blend430": { icon: iconblend430, displayName: "430 Blend" },
  "Home": { icon: iconHome, displayName: "Home" },
  "Tundish": { icon: iconTundish, displayName: "Tundish" },
  "Reclaim": { icon: iconReclaim, displayName: "Reclaim" },
  "Mtown": { icon: iconMtown, displayName: "Middletown" },
  "HBI": { icon: iconHBI, displayName: "Hot Briq Iron" },
  "Frag": { icon: iconFrag, displayName: "Fragmented Scrap" },
  "Alloys": { icon: iconAlloys, displayName: "Alloys" },
  "Other": { icon: iconOther, displayName: "Other" },
};
Object.keys(markerConfig).forEach(type =>
  markerConfig[type].layer = L.layerGroup().addTo(map)
);
 
// ===== Simple accordion coordinator for the two panels =====
window.yardPanels = window.yardPanels || {};
function registerPanel(id, api) { window.yardPanels[id] = api; }
function collapseOthers(exceptId) {
  Object.entries(window.yardPanels).forEach(([id, api]) => {
    if (id !== exceptId && api && typeof api.isExpanded === 'function' &&
        api.isExpanded() && typeof api.collapse === 'function') {
      api.collapse();
    }
  });
}

let allMarkersData = [];
let stockIndexGlobal = {};
let stoppedPileCodesGlobal = new Set();
let searchMode = 'all'; // 'all' | 'pastDue'
let isSearchModalOpen = false;

// Shared styles
const POPUP_CONTAINER_STYLE = 'min-width:420px;max-width:100%;width:auto;';
const ACTIVITY_CONTAINER_STYLE = 'display:none;max-height:320px;overflow-y:auto;overflow-x:hidden;width:auto;box-sizing:border-box;padding:4px;background:#fff';
const ACTIVITY_TABLE_STYLE = 'width:auto;min-width:100%;font-size:12px;border-collapse:collapse';

/* ===================================================================
 ATTENTION "PING" EFFECT
=================================================================== */

function pingMarker(marker, options = {}) {
  if (!marker) return;
  const latlng = marker.getLatLng();
  const {
    color = '#ff3b30',
    pulses = 2,
    duration = 800,
    maxRadius = 40
  } = options;
  
  // Center map on the pinged marker
  map.setView(latlng, 19);
  
  let pulse = 0;
  function makePulse() {
    const circle = L.circle(latlng, {
      radius: 1,
      color,
      weight: 3,
      opacity: 0.9,
      fillColor: color,
      fillOpacity: 0.15
    }).addTo(map);
    const start = performance.now();
    function animate(ts) {
      const t = Math.min(1, (ts - start) / duration);
      const eased = 0.5 - Math.cos(Math.PI * t) / 2;
      circle.setRadius(maxRadius * eased);
      circle.setStyle({ opacity: 0.9 * (1 - t), fillOpacity: 0.25 * (1 - t) });
      if (t < 1) {
        requestAnimationFrame(animate);
      } else {
        map.removeLayer(circle);
        pulse++;
        if (pulse < pulses) setTimeout(makePulse, 120);
      }
    }
    requestAnimationFrame(animate);
  }
  makePulse();
}

/* ===================================================================
 LOAD CELL MARKERS
=================================================================== */
const loadCellMarkers = {
  LC1: L.circleMarker([40.79375707439572, -82.53582134842874], {
    radius: 15, color:"rgba(255,255,0,0.01)", fillColor:"rgba(0,0,0,0.01)",
    fillOpacity:0.01, weight:10
  }).bindTooltip("LC1"),
  LC2: L.circleMarker([40.79414131582521, -82.53631889820099], {
    radius: 15, color:"rgba(255,255,0,0.01)", fillColor:"rgba(0,0,0,0.01)",
    fillOpacity:0.01, weight:10
  }).bindTooltip("LC2"),
  LC3: L.circleMarker([40.79460316783076, -82.53730528056623], {
    radius: 15, color:"rgba(255,255,0,0.01)", fillColor:"rgba(0,0,0,0.01)",
    fillOpacity:0.01, weight:10
  }).bindTooltip("LC3"),
};
Object.values(loadCellMarkers).forEach(m => m.addTo(map));
const loadCells = {
  LC1: L.layerGroup().addTo(map),
  LC2: L.layerGroup().addTo(map),
  LC3: L.layerGroup().addTo(map)
};

/* ===================================================================
 RADIATION LINK MARKER
=================================================================== */
const radiationMarker = L.circleMarker([40.79423553252727, -82.53435252152397], {
  radius: 20,
  color: "rgba(0,0,0,0.00001)",
  fillColor: "rgba(0,0,0,0.00001)",
  fillOpacity: 0.00001,
  weight: 1,
  interactive: true
});
radiationMarker.addTo(map);
radiationMarker.bindTooltip("ASMIV Panel", { permanent: false, direction: 'top', offset: [0, -22] });
radiationMarker.on("click", () => {
  window.open("http://10.141.21.10:8080/ASM/main.jsp", "_blank");
});

/* ===================================================================
 HARSCO SHAREPOINT LINK MARKER
=================================================================== */
const harscoMarker = L.circleMarker([40.79244226436424, -82.53258714613442], {
  radius: 20,
  color: "rgba(0,0,0,0.00001)",
  fillColor: "rgba(0,0,0,0.00001)",
  fillOpacity: 0.00001,
  weight: 1,
  interactive: true
})
  .bindTooltip("Enviri Sharepoint", { permanent: false, direction: 'top' })
  .addTo(map);
harscoMarker.on("click", () => {
  window.open("https://hsconline.sharepoint.com/sites/IX/Pages/IXHome.aspx", "_blank");
});

/* ===================================================================
 LOAD CELL CLICK HANDLER
=================================================================== */
Object.keys(loadCellMarkers).forEach(id => {
  loadCellMarkers[id].on("click", () => {
    if (map.hasLayer(loadCells[id])) {
      map.removeLayer(loadCells[id]);
      return;
    }
    Object.values(markerConfig).forEach(cfg => {
      if (map.hasLayer(cfg.layer)) map.removeLayer(cfg.layer);
    });
    map.addLayer(loadCells[id]);
    map.setView(loadCellMarkers[id].getLatLng(), 20);
  });
});

/* ===================================================================
 TIMESTAMP BANNER
=================================================================== */
(function addTimestampBanner(){
  const ctrl = L.control({ position: 'topright' });
  ctrl.onAdd = function() {
    const div = L.DomUtil.create('div');
    div.tabIndex = 0;
    div.id = 'invBanner';
    div.style.background = 'rgba(255,255,255,0.9)';
    div.style.border = '1px solid #ccc';
    div.style.borderRadius = '4px';
    div.style.padding = '6px 10px';
    div.style.margin = '8px';
    div.style.font = '12px/1.2 system-ui, -apple-system, Segoe UI, Roboto, Arial, sans-serif';
    div.style.boxShadow = '0 1px 3px rgba(0,0,0,0.15)';
    div.innerHTML = '<div>Inventory data current as of —</div><div style="margin-top:4px;color:#555">Total yard inventory: —</div>';
    return div;
  };
  ctrl.addTo(map);
  fetchLatestInventoryCsv().then(payload => {
    const d = payload && payload.meta && payload.meta.report_date;
    const totalInventoryLbs = payload && payload.meta && payload.meta.total_inventory_lbs;
    const banner = document.getElementById('invBanner');
    if (banner) {
      const formattedDate = formatInventoryBannerDate(d);
      const inventoryText = typeof totalInventoryLbs === 'number' && isFinite(totalInventoryLbs)
        ? totalInventoryLbs.toLocaleString('en-US') + ' lbs'
        : '—';
      banner.innerHTML = `
        <div>Inventory data current as of ${formattedDate}</div>
        <div style="margin-top:4px;color:#555">Total yard inventory: <b>${inventoryText}</b></div>
      `;
    }
  }).catch(err => console.warn('stockData.json meta fetch failed:', err));
})();


/* ===================================================================
   BURNING STATION  — stand-alone (not in markers.json / overlay)
   =================================================================== */

// 1) Config -----------------------------------------------------------
const totalsXlsxUrl = 'Production.xlsx';
const burningLatLng = [40.79365495632949, -82.53501357377355];
const breakingLatLng = [40.79408763021845, -82.5388475264302];
const bucketLoadingLatLng = [40.79393364810161, -82.53693358538756];
const receivingLatLng = [40.79379326877625, -82.53696516452395];
const railcarLatLng = [40.79320323425653, -82.53296183080616];

// 2) Cache + helpers --------------------------------------------------
let totalsWorkbookCache = { at: 0, workbook: null };
let burningCache = { at: 0, data: null };
let breakingCache = { at: 0, data: null };
let bucketLoadingCache = { at: 0, data: null };
let receivingCache = { at: 0, data: null };
let railcarCache = { at: 0, data: null };
let historyWorkbookCache = { at: 0, workbook: null };
const receivingHistoryCache = {};
const bucketHistoryCache = {};
let receivingHistoryMonths = null;
let bucketHistoryMonths = null;
let burningHistoryMonths = null;
let breakingHistoryMonths = null;
let railcarHistoryMonths = null;
let receivingCurrentPeriod = null;
let bucketCurrentPeriod = null;
let burningCurrentPeriod = null;
let breakingCurrentPeriod = null;
let railcarCurrentPeriod = null;
const burningHistoryCache = {};
const breakingHistoryCache = {};
const railcarHistoryCache = {};
let railcarPhotoOverlay = null;
let latestInventoryPeriod = { year: null, month: null };

function fmtInt(n)  { return (typeof n === 'number' && isFinite(n)) ? Math.round(n).toLocaleString('en-US') : '—'; }

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

function fmtTons(n, d = 3) { return (typeof n === 'number' && isFinite(n)) ? n.toFixed(d) : '—'; }
function fmtTons2(n, d = 2) { return (typeof n === 'number' && isFinite(n)) ? n.toFixed(d) : '—'; }

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

// ===================== GLOBAL HELPERS (must be above builders) =====================

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

// Last-zero color policy (kept exactly as your current behavior)
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

// 3) CSV loader -------------------------------------------------------
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

async function fetchBurningTotals(force = false, periodOverride = null) {
  const now = Date.now();
  if (periodOverride) {
    const key = `${periodOverride.year}-${periodOverride.month}`;
    if (!force && burningHistoryCache[key]) return burningHistoryCache[key];
  } else {
    if (!force && burningCache.data && (now - burningCache.at) < 120000) {
      return burningCache.data;
    }
  }

  const rows = [];
  const num = v => {
    const x = Number((v || '').toString().trim());
    return Number.isFinite(x) ? x : null;
  };

  const workbook = periodOverride ? await loadHistoryWorkbook(force) : await loadTotalsWorkbook(force);
  const sheetName = periodOverride ? `${periodOverride.year}BurningHistory` : 'Burning';
  const sheet = findSheetByName(workbook, sheetName);
  if (!sheet) throw new Error(`${sheetName} sheet not found`);

  const data = window.XLSX.utils.sheet_to_json(sheet, { header: 1, raw: false });
  if (!Array.isArray(data) || data.length < 2) {
    throw new Error('Burning sheet has no rows');
  }

  const headers = data[0].map(h => String(h || '').trim());
  const template = ['Date','Net(lbs)','From','To','NetTons','# of Cuts','Billable Tons'];
  const blockStarts = findRepeatedBlockStartIndexes(headers, template);

  const fallbackCol = headers.map(h => String(h).toLowerCase());
  const colIndex = names => {
    const lowerNames = names.map(n => String(n).trim().toLowerCase());
    return fallbackCol.findIndex(h => lowerNames.includes(h));
  };

  const fallbackDateCol = colIndex(['date','date/time','datetime']);
  const fallbackNetLbsCol = colIndex(['net(lbs)','net lbs','net','netlbs']);
  const fallbackFromCol = colIndex(['from','from pile #','from pile']);
  const fallbackToCol = colIndex(['to','to pile #','to pile']);
  const fallbackNetTonsCol = colIndex(['nettons','net tons','net_tons']);
  const fallbackCutsCol = colIndex(['# of cuts','cuts','cut']);
  const fallbackBillableCol = colIndex(['billable tons','billabletons','billable']);

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (!Array.isArray(row)) continue;

    if (blockStarts.length > 0) {
      for (const start of blockStarts) {
        const dateCell = row[start];
        const dt = parseXlsxDateCell(dateCell);
        if (!dt) continue;
        const dateLabel = String(dateCell ?? '').trim();

        rows.push({
          date: dt,
          dateLabel,
          netLbs: num(row[start + 1] ?? ''),
          from: String(row[start + 2] || '').trim(),
          to: String(row[start + 3] || '').trim(),
          netTons: num(row[start + 4] ?? ''),
          cuts: num(row[start + 5] ?? ''),
          billableTons: num(row[start + 6] ?? '')
        });
      }
    } else {
      const dateCell = row[fallbackDateCol];
      const dt = parseXlsxDateCell(dateCell);
      if (!dt) continue;
      const dateLabel = String(dateCell ?? '').trim();

      rows.push({
        date: dt,
        dateLabel,
        netLbs: num(row[fallbackNetLbsCol] ?? ''),
        from: String(row[fallbackFromCol] || '').trim(),
        to: String(row[fallbackToCol] || '').trim(),
        netTons: num(row[fallbackNetTonsCol] ?? ''),
        cuts: num(row[fallbackCutsCol] ?? ''),
        billableTons: num(row[fallbackBillableCol] ?? '')
      });
    }
  }

  rows.sort((a, b) => b.date - a.date);

  const { year: filterYear, month: filterMonth } = periodOverride || getCurrentInventoryPeriod();
  const monthRows = rows.filter(r =>
    r.date.getFullYear() === filterYear && r.date.getMonth() === filterMonth
  );

  let totalLbs = 0, totalTons = 0, totalCuts = 0, totalBillable = 0, latest = null;
  for (const r of monthRows) {
    totalLbs += r.netLbs || 0;
    totalTons += r.netTons || 0;
    totalCuts += r.cuts || 0;
    totalBillable += r.billableTons || 0;
    if (!latest || r.date > latest) latest = r.date;
  }

  const payload = {
    rows: monthRows,
    allRows: rows,
    month: {
      year: filterYear,
      monthIndex: filterMonth,
      totalMonthTons: monthRows.reduce((sum, r) => sum + (r.netTons || 0), 0),
      totalMonthLbs: totalLbs,
      rowCount: monthRows.length
    },
    totals: { netLbs: totalLbs, netTons: totalTons, cuts: totalCuts, billableTons: totalBillable },
    latestDateLabel: latest ? latest.toISOString().slice(0,10) : '—'
  };

  window.currentBurningSheetRows = monthRows;
  if (periodOverride) {
    burningHistoryCache[`${periodOverride.year}-${periodOverride.month}`] = payload;
  } else {
    burningCache = { at: now, data: payload };
  }
  return payload;
}

async function fetchBreakingTotals(force = false, periodOverride = null) {
  const now = Date.now();
  if (periodOverride) {
    const key = `${periodOverride.year}-${periodOverride.month}`;
    if (!force && breakingHistoryCache[key]) return breakingHistoryCache[key];
  } else {
    if (!force && breakingCache.data && (now - breakingCache.at) < 120000) {
      return breakingCache.data;
    }
  }

  const rows = [];
  const toNum = (v) => {
    const x = Number((v || '').toString().trim());
    return Number.isFinite(x) ? x : null;
  };

  const workbook = periodOverride ? await loadHistoryWorkbook(force) : await loadTotalsWorkbook(force);
  const sheetName = periodOverride ? `${periodOverride.year}BreakingHistory` : 'Breaking';
  const sheet = findSheetByName(workbook, sheetName);
  if (!sheet) throw new Error(`${sheetName} sheet not found`);

  const data = window.XLSX.utils.sheet_to_json(sheet, { header: 1, raw: false });
  if (!Array.isArray(data) || data.length < 2) {
    throw new Error('Breaking sheet has no rows');
  }

  const headers = data[0].map(h => String(h || '').trim());
  const template = ['Date','Commodity','Net','Net Tons','From Pile #','To Pile #'];
  const blockStarts = findRepeatedBlockStartIndexes(headers, template);

  const fallbackCol = headers.map(h => String(h).toLowerCase());
  const colIndex = names => {
    const lowerNames = names.map(n => String(n).trim().toLowerCase());
    return fallbackCol.findIndex(h => lowerNames.includes(h));
  };

  const dateCol = colIndex(['date','date/time','datetime']);
  const commodityCol = colIndex(['commodity','material']);
  const netLbsCol = colIndex(['net','net lbs','netlbs']);
  const netTonsCol = colIndex(['net tons','nettons','net_tons']);
  const fromCol = colIndex(['from','from pile #','from pile']);
  const toCol = colIndex(['to','to pile #','to pile']);

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (!Array.isArray(row)) continue;

    if (blockStarts.length > 0) {
      for (const start of blockStarts) {
        const dateCell = row[start];
        const dt = parseXlsxDateCell(dateCell);
        if (!dt) continue;
        const dateLabel = String(dateCell ?? '').trim();

        rows.push({
          date: dt,
          dateLabel,
          material: String(row[start + 1] || '').trim(),
          netLbs: toNum(row[start + 2] ?? ''),
          netTons: toNum(row[start + 3] ?? ''),
          from: String(row[start + 4] || '').trim(),
          to: String(row[start + 5] || '').trim()
        });
      }
    } else {
      const dateCell = row[dateCol];
      const dt = parseXlsxDateCell(dateCell);
      if (!dt) continue;

      const dateLabel = String(dateCell ?? '').trim();
      rows.push({
        date: dt,
        dateLabel,
        material: String(row[commodityCol] || '').trim(),
        netLbs: toNum(row[netLbsCol] ?? ''),
        netTons: toNum(row[netTonsCol] ?? ''),
        from: String(row[fromCol] || '').trim(),
        to: String(row[toCol] || '').trim()
      });
    }
  }

  rows.sort((a, b) => b.date - a.date);

  const { year: filterYear, month: filterMonth } = periodOverride || getCurrentInventoryPeriod();
  const monthRows = rows.filter(r =>
    r.date.getFullYear() === filterYear && r.date.getMonth() === filterMonth
  );

  let totalLbs = 0;
  let totalTons = 0;
  for (const r of monthRows) {
    totalLbs += r.netLbs || 0;
    totalTons += r.netTons || 0;
  }

  const payload = {
    rows: monthRows,
    allRows: rows,
    month: {
      year: filterYear,
      monthIndex: filterMonth,
      totalMonthTons: totalTons,
      totalMonthLbs: totalLbs,
      rowCount: monthRows.length
    },
    totals: {
      totalLbs,
      totalTons
    }
  };

  window.currentBreakingSheetRows = monthRows;
  if (periodOverride) {
    breakingHistoryCache[`${periodOverride.year}-${periodOverride.month}`] = payload;
  } else {
    breakingCache = { at: now, data: payload };
  }
  return payload;
}

// 4) Popup rendering
function buildUnprepRows(markers, stockIndex) {
  return markers
    .filter(m => m.type === "Breaking" || m.type === "Unbreakable")
    .map(m => {
      const code = extractPileCode(m.name);
      const s = (code && stockIndex[code]) ? stockIndex[code] : {};
      const inv = (typeof s.operating_inventory_lbs === 'number')
        ? s.operating_inventory_lbs
        : Number(s.operating_inventory_lbs || 0);
      const lastZero = s.last_zero_date ?? '—';

      const invStyle = invAlertStyle(inv);
      const lzStyle  = lastZeroColor(lastZero);

      return `
        <tr>
          <td style="padding:2px 6px">${m.name}</td>
          <td style="padding:2px 6px;text-align:right;${invStyle}">${Number.isFinite(inv) ? inv.toLocaleString('en-US') : '—'}</td>
          <td style="padding:2px 6px;${lzStyle}">${lastZero}</td>
        </tr>
      `;
    })
    .join('');
}

function buildCoilsRows(markers, stockIndex) {
  return markers
    .filter(m => m.type === "Coils")
    .map(m => {
      const code = extractPileCode(m.name);
      const s = code ? (stockIndex[code] || {}) : {};
      const inv = (typeof s.operating_inventory_lbs === 'number')
        ? s.operating_inventory_lbs
        : Number(s.operating_inventory_lbs || 0);
      const lastZero = s.last_zero_date ?? '—';

      const invStyle = invAlertStyle(inv);
      const lzStyle  = lastZeroColor(lastZero);

      return `
        <tr>
          <td style="padding:2px 6px">${m.name}</td>
          <td style="padding:2px 6px;text-align:right;${invStyle}">${Number.isFinite(inv) ? inv.toLocaleString('en-US') : '—'}</td>
          <td style="padding:2px 6px;${lzStyle}">${lastZero}</td>
        </tr>
      `;
    })
    .join('');
}

function getUnprepTotalInventory(markers, stockIndex) {
  let total = 0;
  for (const m of markers) {
    if (m.type !== "Breaking" && m.type !== "Unbreakable") continue;
    const code = extractPileCode(m.name);
    if (!code) continue;
    const s = stockIndex[code];
    const v = (s && typeof s.operating_inventory_lbs === 'number') ? s.operating_inventory_lbs : 0;
    total += v;
  }
  return total;
}

function getCoilsTotalInventory(markers, stockIndex) {
  let total = 0;
  for (const m of markers) {
    if (m.type !== "Coils") continue;
    const code = extractPileCode(m.name);
    if (!code) continue;
    const s = stockIndex[code];
    const v = (s && typeof s.operating_inventory_lbs === 'number') ? s.operating_inventory_lbs : 0;
    total += v;
  }
  return total;
}

function renderBreakingPopup(payload, markers, stockIndex) {
  if (!payload) {
    return '<b>Breaking Pit</b><div>No data.</div>';
  }

  const monthSelectorHtml = buildMonthSelectorHtml('breaking', breakingCurrentPeriod, breakingHistoryMonths);

  // --- Breaking summary (Unprep first, then Processed; no mismatched <b>) ---
const t = payload.totals || {};
const totalProcessedLbs  = (typeof t.totalLbs  === 'number' && isFinite(t.totalLbs))   ? t.totalLbs  : null;
const totalProcessedTons = (typeof t.totalTons === 'number' && isFinite(t.totalTons)) ? t.totalTons : null;
const totalProcessedText = (totalProcessedLbs !== null)
  ? `${fmtInt(totalProcessedLbs)} <span style="color:#555">(${fmtTons2(totalProcessedTons ?? 0, 2)} tons)</span>`
  : '—';

const unprepTotalLbs = getUnprepTotalInventory(markers, stockIndex);
const unprepLbsText  = isFinite(unprepTotalLbs) ? unprepTotalLbs.toLocaleString('en-US') : '—';
const unprepTonsText = isFinite(unprepTotalLbs) ? (unprepTotalLbs / 2000).toFixed(2) : '—';

const summary = `
  <div style="margin-bottom:8px; text-align:center">
    <table style="width:auto;font-size:12px;line-height:1.3;border-collapse:collapse;margin:0 auto;text-align:left">
      <tr>
        <td style="color:#666;padding:2px 6px">Total Unprep Inventory</td>
        <td style="text-align:right;padding:2px 6px">
          <b>${unprepLbsText}</b> <span style="color:#555"><b>(${unprepTonsText} tons)</b></span>
        </td>
      </tr>
      <tr>
        <td style="color:#666;padding:2px 6px">Total Processed</td>
        <td style="text-align:right;padding:2px 6px">
          <b>${totalProcessedText}</b>
        </td>
      </tr>
    </table>
  </div>
`;

// Activity rows - ALL for current month
  const rowsHtml = (payload.rows || []).map(r => `
    <tr>
      <td style="padding:2px 6px;white-space:nowrap">${renderDateWithWeekday(r.date, r.dateLabel)}</td>
      <td style="padding:2px 6px;text-align:center;white-space:nowrap">${esc(r.from)}</td>
      <td style="padding:2px 6px;text-align:center;white-space:nowrap">→</td>
      <td style="padding:2px 6px;white-space:nowrap">${esc(r.to)}</td>
      <td style="padding:2px 6px;text-align:right">${fmtTons2(r.netTons, 3)}</td>
      <td style="padding:2px 6px">${esc(r.material || '')}</td>
    </tr>
  `).join('');

  const rowCount = (payload.rows || []).length;
  const showCountText = rowCount > 0 ? `<div style="font-size:11px;color:#999;margin:4px 0;">${rowCount} transfers this month</div>` : '';
  
  // Store totals globally for CSV export
  window.currentBreakingTotals = payload.totals || {};

  
const activity = `
  <div id="breakingActivity" style="${ACTIVITY_CONTAINER_STYLE}">
    ${showCountText}
    <table style="${ACTIVITY_TABLE_STYLE}">
      <thead>
        <tr style="background:#f2f2f2">
          <th style="text-align:left;padding:2px 6px">Date</th>
          <th style="text-align:center;padding:2px 6px">From</th>
          <th style="text-align:center;padding:2px 6px">→</th>
          <th style="text-align:left;padding:2px 6px">To</th>
          <th style="text-align:right;padding:2px 6px">Net Tons</th>
          <th style="text-align:left;padding:2px 6px">Material</th>
        </tr>
      </thead>
      <tbody>${rowsHtml}</tbody>
    </table>
  </div>
  <div style="margin-top:6px; display:flex; gap:6px; align-items:center">
    <button type="button" id="breakingToggle"
      style="padding:2px 6px;border:1px solid #ddd;border-radius:3px;background:#f8f8f8;cursor:pointer">
      Show Activity
    </button>

    <!-- NEW: mirror Burning's Download CSV button -->
    <button type="button" id="breakingDownload"
      style="padding:2px 6px;border:1px solid #ddd;border-radius:3px;background:#f8f8f8;cursor:pointer; display:none">
      Download CSV
    </button>

    <!-- Top Unprep toggle -->
    <button type="button" id="unprepToggleTop"
      style="padding:2px 6px;border:1px solid #ddd;border-radius:3px;background:#f8f8f8;cursor:pointer">
      Show Unprep
    </button>
  </div>
`;

  // Unprep section + bottom button (mirrors coils section)
  const unprepRowsHtml = buildUnprepRows(markers, stockIndex);
  const unprepSection = `
  <div id="unprepSection" style="display:none;margin-top:8px">
    <table style="width:100%;font-size:12px;border-collapse:collapse">
      <thead>
        <tr style="background:#f2f2f2">
          <th style="text-align:left;padding:2px 6px">Pile</th>
          <th style="text-align:right;padding:2px 6px">Inventory (lbs)</th>
          <th style="text-align:left;padding:2px 6px">Last Zero</th>
        </tr>
      </thead>
      <tbody>
        ${unprepRowsHtml}
      </tbody>
    </table>
    <div style="margin-top:6px; display:flex; justify-content:flex-end">
      <button type="button" id="unprepToggleBottom"
        style="padding:2px 6px;border:1px solid #ddd;border-radius:3px;background:#f8f8f8;cursor:pointer">
        Hide Unprep
      </button>
    </div>
  </div>
  `;

  const body = `
    <div style="margin-bottom:4px">${monthSelectorHtml}</div>
    <div style="font-weight:700;margin-bottom:6px">Breaking Pit</div>
    ${summary}
    ${activity}
    ${unprepSection}
  `;
  return `<div style="${POPUP_CONTAINER_STYLE}">${body}</div>`;
}

function renderBurningPopup(payload, markers, stockIndex) {
  if (!payload) {
    return '&lt;b&gt;Burning Station&lt;/b&gt;&lt;div&gt;No data.&lt;/div&gt;';
  }

  const monthSelectorHtml = buildMonthSelectorHtml('burning', burningCurrentPeriod, burningHistoryMonths);

  const t = payload.totals;
  const coilsRowsHtml = buildCoilsRows(markers, stockIndex);
  const coilsInvLbs = getCoilsTotalInventory(markers, stockIndex);
  const coilsInvText = (typeof coilsInvLbs === 'number' && isFinite(coilsInvLbs))
    ? coilsInvLbs.toLocaleString('en-US')
    : '—';

  // Show tons to two decimals if we have a numeric value
  const coilsInvTonsText = (typeof coilsInvLbs === 'number' && isFinite(coilsInvLbs))
    ? (coilsInvLbs / 2000).toFixed(2)
    : '—';

  // Summary section

const summary = `
  <div style="margin-bottom:8px;text-align:center">
    <table style="width:auto;font-size:12px;line-height:1.3;border-collapse:collapse;margin:0 auto; text-align:left">
      <tr>
        <td style="color:#666;padding:2px 6px">Total Coil Inventory</td>
        <td style="text-align:left;padding:2px 6px"><b>${coilsInvText}</b> <span style="color:#555"><b>(${coilsInvTonsText} tons)</b></span></td>
      </tr>
      <tr>
        <td style="color:#666;padding:2px 6px">Net Tons Cut</td>
        <td style="text-align:left;padding:2px 6px"><b>${fmtTons(t.netTons)}</b></td>
      </tr>
      <tr>
        <td style="color:#666;padding:2px 6px">Billable Tons Cut</td>
        <td style="text-align:left;padding:2px 6px"><b>${fmtTons(t.billableTons)}</b></td>
      </tr>
    </table>
  </div>
`;

  // Activity rows - ALL for current month
  const rowsHtml = (payload.rows || []).map(r => `
    &lt;tr&gt;
      &lt;td style="padding:2px 6px;white-space:nowrap"&gt;${renderDateWithWeekday(r.date, r.dateLabel)}&lt;/td&gt;
      &lt;td style="padding:2px 6px;text-align:center;white-space:nowrap"&gt;${esc(r.from)}&lt;/td&gt;
      &lt;td style="padding:2px 6px;text-align:center;white-space:nowrap"&gt;→&lt;/td&gt;
      &lt;td style="padding:2px 6px;white-space:nowrap"&gt;${esc(r.to)}&lt;/td&gt;
      &lt;td style="padding:2px 6px;text-align:right"&gt;${fmtTons(r.netTons)}&lt;/td&gt;
      &lt;td style="padding:2px 6px;text-align:right"&gt;${fmtInt(r.cuts)}&lt;/td&gt;
      &lt;td style="padding:2px 6px;text-align:right"&gt;${fmtTons(r.billableTons)}&lt;/td&gt;
    &lt;/tr&gt;
  `).join('');

  const totalTransfers = (payload.rows || []).length;
  const activityCountHint = `<div style="font-size:11px;color:#999;margin:4px 0;">${totalTransfers} transfers this month</div>`;
  
  // Store totals globally for CSV export
  window.currentBurningTotals = payload.totals || {};

  // Activity section
  const activity = `
  &lt;div id="burningActivity" style="${ACTIVITY_CONTAINER_STYLE}"&gt;
      ${activityCountHint}
      &lt;table style="${ACTIVITY_TABLE_STYLE}"&gt;
        &lt;thead&gt;
          &lt;tr style="background:#f2f2f2"&gt;
            &lt;th style="text-align:left;padding:2px 6px"&gt;Date&lt;/th&gt;
            &lt;th style="text-align:left;padding:2px 6px"&gt;From&lt;/th&gt;
            &lt;th style="text-align:center;padding:2px 6px"&gt;→&lt;/th&gt;
            &lt;th style="text-align:left;padding:2px 6px"&gt;To&lt;/th&gt;
            &lt;th style="text-align:right;padding:2px 6px"&gt;Net Tons&lt;/th&gt;
            &lt;th style="text-align:right;padding:2px 6px"&gt;Cuts&lt;/th&gt;
            &lt;th style="text-align:right;padding:2px 6px"&gt;Billable&lt;/th&gt;
          &lt;/tr&gt;
        &lt;/thead&gt;
        &lt;tbody&gt;${rowsHtml}&lt;/tbody&gt;
      &lt;/table&gt;
    &lt;/div&gt;

    &lt;div style="margin-top:6px; display:flex; gap:6px; align-items:center"&gt;
      &lt;button type="button" id="burningToggle"
        style="padding:2px 6px;border:1px solid #ddd;border-radius:3px;background:#f8f8f8;cursor:pointer"&gt;
        Show Activity
      &lt;/button&gt;

      &lt;button type="button" id="burningDownload"
        style="padding:2px 6px;border:1px solid #ddd;border-radius:3px;background:#f8f8f8;cursor:pointer; display:none"&gt;
        Download CSV
      &lt;/button&gt;

      &lt;button type="button" id="coilsToggleTop"
        style="padding:2px 6px;border:1px solid #ddd;border-radius:3px;background:#f8f8f8;cursor:pointer"&gt;
        Show Coils
      &lt;/button&gt;
  &lt;/div&gt;
  `;

  const coilsSection = `
  <div id="coilsSection" style="display:none;margin-top:8px">
    <table style="width:100%;font-size:12px;border-collapse:collapse">
      <thead>
        <tr style="background:#f2f2f2">
          <th style="text-align:left;padding:2px 6px">Pile</th>
          <th style="text-align:right;padding:2px 6px">Inventory (lbs)</th>
          <th style="text-align:left;padding:2px 6px">Last Zero</th>
        </tr>
      </thead>
      <tbody>
        ${coilsRowsHtml}
      </tbody>
    </table>
    <div style="margin-top:6px; display:flex; justify-content:flex-end">
      <button type="button" id="coilsToggleBottom"
        style="padding:2px 6px;border:1px solid #ddd;border-radius:3px;background:#f8f8f8;cursor:pointer">
        Hide Coils
      </button>
    </div>
  </div>
`;


  const body = `
    <div style="margin-bottom:4px">${monthSelectorHtml}</div>
    &lt;div style="font-weight:700;margin-bottom:6px"&gt;Burning Station&lt;/div&gt;
    ${summary}
    ${activity}
    ${coilsSection}
  `;

  return `&lt;div style="${POPUP_CONTAINER_STYLE}"&gt;${body}&lt;/div&gt;`;
}

// 5) Wire popup events ------------------------------------------------

async function fetchBucketLoadingConsumption(force = false, periodOverride = null) {
  const now = Date.now();
  if (periodOverride) {
    const key = `${periodOverride.year}-${periodOverride.month}`;
    if (!force && bucketHistoryCache[key]) {
      window.currentBucketLoadingRows = bucketHistoryCache[key].rows;
      return bucketHistoryCache[key];
    }
  } else {
    if (!force && bucketLoadingCache.data && (now - bucketLoadingCache.at) < 120000) {
      return bucketLoadingCache.data;
    }
  }

  const toNum = v => {
    if (v == null) return 0;
    if (typeof v === 'number') return Number.isFinite(v) ? v : 0;

    const raw = String(v).trim();
    if (!raw) return 0;

    // Handle common Excel/export formats like "1,234", "(1,234)", or "1,234 lbs".
    const cleaned = raw
      .replace(/,/g, '')
      .replace(/\(([^)]+)\)/, '-$1')
      .replace(/[^0-9.+-]/g, '');

    if (!cleaned || cleaned === '-' || cleaned === '+') return 0;
    const n = Number(cleaned);
    return Number.isFinite(n) ? n : 0;
  };

  const workbook = periodOverride ? await loadHistoryWorkbook(force) : await loadTotalsWorkbook(force);
  const sheet = findSheetByName(workbook, historySheetName('Consumption1', periodOverride));
  if (!sheet) {
    console.warn('Consumption1 sheet not found');
    return null;
  }

  const data = window.XLSX.utils.sheet_to_json(sheet, { header: 1, raw: false });
  if (!Array.isArray(data) || data.length < 2) {
    console.warn('Consumption1 sheet has insufficient data');
    return null;
  }

  const { year: filterYear, month: filterMonth } = periodOverride || getCurrentInventoryPeriod();

  const monthMatches = (date) => {
    return date instanceof Date && !Number.isNaN(date.getTime()) && date.getFullYear() === filterYear && date.getMonth() === filterMonth;
  };

  const toIsoFromDate = (dt) => dt instanceof Date && !Number.isNaN(dt.getTime())
    ? `${dt.getFullYear()}-${String(dt.getMonth() + 1).padStart(2, '0')}-${String(dt.getDate()).padStart(2, '0')}`
    : '';

  const headerRow = Array.isArray(data[0]) ? data[0] : [];
  const headerNorm = headerRow.map(h => String(h || '').trim().toLowerCase());
  const findCol = (...names) => {
    for (const name of names) {
      const idx = headerNorm.findIndex(h => h === String(name).toLowerCase());
      if (idx >= 0) return idx;
    }
    return -1;
  };

  const dateCol = findCol('date');
  const poundsCol = findCol('total lbs', 'lbs', 'pounds');
  const tonsCol = findCol('total tons', 'tons');
  const bucketsCol = findCol('buckets loaded', 'buckets');
  const heatsCol = findCol('heats completed', 'heats');

  const resolvedDateCol = dateCol >= 0 ? dateCol : 0;
  const resolvedPoundsCol = poundsCol >= 0 ? poundsCol : 1;
  const resolvedTonsCol = tonsCol >= 0 ? tonsCol : 2;
  const resolvedBucketsCol = bucketsCol >= 0 ? bucketsCol : 3;
  const resolvedHeatsCol = heatsCol >= 0 ? heatsCol : 4;

  const breakdownByIsoDate = {};

  const normalizeKey = (v) => String(v ?? '').trim().toUpperCase();
  const getMaterialForLotOrPile = (lot, pile) => {
    const lotKey = normalizeKey(lot);
    const pileKey = normalizePileCode(pile);

    if (lotKey) {
      const byLot = Object.values(stockIndexGlobal || {}).find(item =>
        normalizeKey(item && item.code) === lotKey
      );
      if (byLot && byLot.material) return String(byLot.material);
    }

    if (pileKey && stockIndexGlobal && stockIndexGlobal[pileKey] && stockIndexGlobal[pileKey].material) {
      return String(stockIndexGlobal[pileKey].material);
    }

    return '';
  };

  const consumption2 = findSheetByName(workbook, historySheetName('Consumption2', periodOverride));
  if (consumption2) {
    const data2 = window.XLSX.utils.sheet_to_json(consumption2, { header: 1, raw: false });
    if (Array.isArray(data2) && data2.length >= 2) {
      const header2 = Array.isArray(data2[0]) ? data2[0] : [];
      const header2Norm = header2.map(h => String(h || '').trim().toLowerCase());
      const findCol2 = (...names) => {
        for (const name of names) {
          const idx = header2Norm.findIndex(h => h === String(name).toLowerCase());
          if (idx >= 0) return idx;
        }
        return -1;
      };

      const date2Col = (() => {
        const i = findCol2('date', 'date material was used');
        return i >= 0 ? i : 0;
      })();
      const pile2Col = (() => {
        const i = findCol2('pile utilized', 'pile used', 'pile', 'pile #');
        return i >= 0 ? i : 1;
      })();
      const lot2Col = (() => {
        const i = findCol2('material lot# utilized', 'material lot # utilized', 'material lot # used', 'material lot #', 'lot #', 'lot');
        return i >= 0 ? i : 2;
      })();
      const pounds2Col = (() => {
        const i = findCol2('pounds utilized', 'total weight of material used in pounds', 'total weight of material used', 'weight (lbs)', 'pounds', 'lbs');
        return i >= 0 ? i : 3;
      })();
      const tons2Col = (() => {
        const i = findCol2('tons utilized', 'tons used', 'tons');
        return i >= 0 ? i : 4;
      })();

      const pileBreakdownByDate = {};

      for (let r2 = 1; r2 < data2.length; r2++) {
        const row2 = data2[r2];
        if (!Array.isArray(row2)) continue;
        const dateCell2 = row2[date2Col];
        const parsed2 = parseXlsxDateCell(dateCell2);

        let iso2 = '';
        if (parsed2) {
          iso2 = toIsoFromDate(parsed2);
        } else {
          const maybeDay2 = Number(String(dateCell2 || '').trim());
          if (Number.isFinite(maybeDay2) && maybeDay2 >= 1 && maybeDay2 <= 31) {
            iso2 = `${filterYear}-${String(filterMonth + 1).padStart(2, '0')}-${String(maybeDay2).padStart(2, '0')}`;
          }
        }

        if (!iso2) continue;

        const pile = normalizePileCode(row2[pile2Col]);
        const lot = String(row2[lot2Col] || '').trim();
        const pounds = toNum(row2[pounds2Col]);
        const tonsCellRaw = row2[tons2Col];
        const tonsCellText = String(tonsCellRaw ?? '').trim();
        const rawTons = toNum(tonsCellRaw);
        const tons = tonsCellText === '' ? (pounds / 2000) : rawTons;

        if (!pile || !Number.isFinite(pounds) || pounds === 0) continue;

        if (!pileBreakdownByDate[iso2]) {
          pileBreakdownByDate[iso2] = {};
        }

        if (!pileBreakdownByDate[iso2][pile]) {
          pileBreakdownByDate[iso2][pile] = { pile, pounds: 0, tons: 0, lots: new Map() };
        }

        pileBreakdownByDate[iso2][pile].pounds += pounds;
        pileBreakdownByDate[iso2][pile].tons += tons;
        if (lot) {
          const material = getMaterialForLotOrPile(lot, pile);
          pileBreakdownByDate[iso2][pile].lots.set(lot, material);
        }
      }

      Object.keys(pileBreakdownByDate).forEach(iso => {
        const pileMap = pileBreakdownByDate[iso];
        breakdownByIsoDate[iso] = Object.values(pileMap)
          .map(item => ({
            pile: item.pile,
            pounds: item.pounds,
            tons: item.tons,
            lotCount: item.lots.size,
            lots: Array.from(item.lots.entries()).map(([lot, material]) => ({ lot, material }))
          }))
          .sort((a, b) => b.pounds - a.pounds || a.pile.localeCompare(b.pile));
      });
    }
  }

  // ---- Consumption4: heat-level material breakdown ----
  const heatBreakdownByIsoDate = {};
  const consumption4 = findSheetByName(workbook, historySheetName('Consumption4', periodOverride));
  if (consumption4) {
    const data4 = window.XLSX.utils.sheet_to_json(consumption4, { header: 1, raw: false });
    if (Array.isArray(data4) && data4.length >= 2) {
      const header4 = Array.isArray(data4[0]) ? data4[0] : [];
      const header4Norm = header4.map(h => String(h || '').trim().toLowerCase());
      const findCol4 = (...names) => {
        for (const name of names) {
          const idx = header4Norm.findIndex(h => h === String(name).toLowerCase());
          if (idx >= 0) return idx;
        }
        return -1;
      };

      const date4Col  = (() => { const i = findCol4('date'); return i >= 0 ? i : 0; })();
      const heat4Col  = (() => { const i = findCol4('heat number', 'heat #', 'heat'); return i >= 0 ? i : 1; })();
      const grade4Col = (() => { const i = findCol4('grade'); return i >= 0 ? i : 2; })();
      const pile4Col  = (() => { const i = findCol4('pile', 'pile #', 'pile utilized'); return i >= 0 ? i : 3; })();
      const lot4Col   = (() => { const i = findCol4('material lot #', 'lot #', 'lot'); return i >= 0 ? i : 4; })();
      const lbs4Col   = (() => { const i = findCol4('total pounds', 'pounds', 'lbs'); return i >= 0 ? i : 5; })();
      const tons4Col  = (() => { const i = findCol4('total tons', 'tons'); return i >= 0 ? i : 6; })();

      // Build per-heat, per-day buckets first, then keep each heat only on its latest day.
      // This prevents carryover heats from showing on both start day and completion day.
      const heatBucketsByNumber = {};

      for (let r4 = 1; r4 < data4.length; r4++) {
        const row4 = data4[r4];
        if (!Array.isArray(row4)) continue;

        const dateCell4 = row4[date4Col];
        const parsed4 = parseXlsxDateCell(dateCell4);
        let iso4 = '';
        if (parsed4) {
          iso4 = toIsoFromDate(parsed4);
        } else {
          const maybeDay4 = Number(String(dateCell4 || '').trim());
          if (Number.isFinite(maybeDay4) && maybeDay4 >= 1 && maybeDay4 <= 31) {
            iso4 = `${filterYear}-${String(filterMonth + 1).padStart(2, '0')}-${String(maybeDay4).padStart(2, '0')}`;
          }
        }
        if (!iso4) continue;

        const heatNum4 = String(row4[heat4Col] || '').trim();
        const grade4   = String(row4[grade4Col] || '').trim();
        const pile4    = normalizePileCode(row4[pile4Col]);
        const lot4     = String(row4[lot4Col] || '').trim();
        const lbs4     = toNum(row4[lbs4Col]);
        const tonsCellRaw4 = row4[tons4Col];
        const tons4    = String(tonsCellRaw4 ?? '').trim() === '' ? (lbs4 / 2000) : toNum(tonsCellRaw4);

        if (!heatNum4) continue;

        if (!heatBucketsByNumber[heatNum4]) {
          heatBucketsByNumber[heatNum4] = {};
        }
        if (!heatBucketsByNumber[heatNum4][iso4]) {
          heatBucketsByNumber[heatNum4][iso4] = { heatNumber: heatNum4, grade: grade4, materials: [] };
        } else if (grade4 && !heatBucketsByNumber[heatNum4][iso4].grade) {
          heatBucketsByNumber[heatNum4][iso4].grade = grade4;
        }

        if (pile4 || lot4 || lbs4) {
          const material4 = getMaterialForLotOrPile(lot4, pile4);
          heatBucketsByNumber[heatNum4][iso4].materials.push({ pile: pile4, lot: lot4, material: material4, pounds: lbs4, tons: tons4 });
        }
      }

      Object.keys(heatBucketsByNumber).forEach(heatNum => {
        const byIso = heatBucketsByNumber[heatNum];
        const isoKeys = Object.keys(byIso);
        if (isoKeys.length === 0) return;

        // ISO date strings sort chronologically (YYYY-MM-DD).
        const latestIso = isoKeys.sort().slice(-1)[0];
        if (!heatBreakdownByIsoDate[latestIso]) heatBreakdownByIsoDate[latestIso] = {};

        const selected = byIso[latestIso];
        if (!heatBreakdownByIsoDate[latestIso][heatNum]) {
          heatBreakdownByIsoDate[latestIso][heatNum] = {
            heatNumber: selected.heatNumber,
            grade: selected.grade,
            materials: []
          };
        }

        heatBreakdownByIsoDate[latestIso][heatNum].materials.push(...(selected.materials || []));
        if (!heatBreakdownByIsoDate[latestIso][heatNum].grade && selected.grade) {
          heatBreakdownByIsoDate[latestIso][heatNum].grade = selected.grade;
        }
      });

      Object.keys(heatBreakdownByIsoDate).forEach(iso => {
        const heatMap = heatBreakdownByIsoDate[iso];
        heatBreakdownByIsoDate[iso] = Object.values(heatMap).sort((a, b) => {
          const na = Number(a.heatNumber), nb = Number(b.heatNumber);
          if (!isNaN(na) && !isNaN(nb)) return na - nb;
          return String(a.heatNumber).localeCompare(String(b.heatNumber));
        });
      });
    }
  }

  const rows = [];
  const allRows = [];
  let totalPounds = 0;
  let totalTons = 0;
  let avgPounds = 0;
  let avgTons = 0;
  let totalRowCount = 0;

  for (let r = 1; r < data.length; r++) {
    const rowData = data[r];
    if (!Array.isArray(rowData)) continue;

    const dateCell = rowData[resolvedDateCol];
    const poundsCell = rowData[resolvedPoundsCol];
    const tonsCell = rowData[resolvedTonsCol];
    const bucketsCell = rowData[resolvedBucketsCol];
    const heatsCell = rowData[resolvedHeatsCol];

    if ((dateCell == null || String(dateCell).trim() === '') && (poundsCell == null || String(poundsCell).trim() === '') && (tonsCell == null || String(tonsCell).trim() === '') && (bucketsCell == null || String(bucketsCell).trim() === '') && (heatsCell == null || String(heatsCell).trim() === '')) {
      continue;
    }

    const dateText = String(dateCell || '').trim();
    const lbs = toNum(poundsCell);
    const tons = toNum(tonsCell);
    const bucketsLoaded = toNum(bucketsCell);
    const heatsCompleted = toNum(heatsCell);
    const lowerDateText = dateText.toLowerCase();

    if (lowerDateText.includes('total') || lowerDateText.includes('sum') || lowerDateText.includes('average') || lowerDateText.includes('avg')) {
      if ((lowerDateText.includes('total') || lowerDateText.includes('sum')) && lbs > 0) {
        totalPounds = lbs;
        totalTons = tons;
      }
      if ((lowerDateText.includes('average') || lowerDateText.includes('avg')) && lbs > 0) {
        avgPounds = lbs;
        avgTons = tons;
      }
      continue;
    }

    let dayNum = null;
    let dateLabel = '';
    let isoDate = '';

    const parsed = parseXlsxDateCell(dateCell);
    if (monthMatches(parsed)) {
      dayNum = parsed.getDate();
      dateLabel = `${parsed.getMonth() + 1}/${parsed.getDate()}`;
      isoDate = toIsoFromDate(parsed);
    } else {
      const maybeDay = Number(dateText);
      if (Number.isFinite(maybeDay) && maybeDay >= 1 && maybeDay <= 31) {
        dayNum = maybeDay;
        dateLabel = `Day ${maybeDay}`;
        isoDate = `${filterYear}-${String(filterMonth + 1).padStart(2, '0')}-${String(maybeDay).padStart(2, '0')}`;
      } else {
        continue;
      }
    }

    const heatBreakdownForDate = heatBreakdownByIsoDate[isoDate] || [];
    const resolvedHeatsCompleted = heatBreakdownForDate.length > 0
      ? heatBreakdownForDate.length
      : heatsCompleted;

    const rowObj = {
      day: dayNum,
      dateLabel,
      isoDate,
      pounds: lbs,
      tons,
      bucketsLoaded,
      heatsCompleted: resolvedHeatsCompleted,
      breakdown: breakdownByIsoDate[isoDate] || [],
      heatBreakdown: heatBreakdownForDate
    };

    allRows.push(rowObj);
    if (monthMatches(parsed) || !parsed) {
      rows.push(rowObj);
    }
    totalRowCount += 1;
  }

  // If no rows matched the current month filter, fall back to all valid rows.
  if (rows.length === 0 && allRows.length > 0) {
    rows.push(...allRows);
  }

  if (totalPounds === 0 && totalRowCount > 0) {
    totalPounds = rows.reduce((sum, row) => sum + (row.pounds || 0), 0);
    totalTons = rows.reduce((sum, row) => sum + (row.tons || 0), 0);
  }
  if (avgPounds === 0 && totalRowCount > 0) {
    avgPounds = totalPounds / totalRowCount;
    avgTons = totalTons / totalRowCount;
  }

  const payload = {
    rows,
    month: {
      year: filterYear,
      monthIndex: filterMonth,
      rowCount: totalRowCount,
      breakdownRowCount: Object.keys(breakdownByIsoDate).length
    },
    totals: {
      totalPounds,
      totalTons,
      avgPounds,
      avgTons
    }
  };

  window.currentBucketLoadingRows = rows;
  if (periodOverride) {
    bucketHistoryCache[`${periodOverride.year}-${periodOverride.month}`] = payload;
  } else {
    bucketLoadingCache = { at: now, data: payload };
  }
  return payload;
}

async function fetchReceivingSummary(force = false, periodOverride = null) {
  const now = Date.now();
  if (periodOverride) {
    const key = `${periodOverride.year}-${periodOverride.month}`;
    if (!force && receivingHistoryCache[key]) {
      window.currentReceivingRows = receivingHistoryCache[key].rows;
      return receivingHistoryCache[key];
    }
  } else {
    if (!force && receivingCache.data && (now - receivingCache.at) < 120000) {
      return receivingCache.data;
    }
  }

  const toNum = v => {
    if (v == null) return 0;
    if (typeof v === 'number') return Number.isFinite(v) ? v : 0;

    const raw = String(v).trim();
    if (!raw) return 0;

    const cleaned = raw
      .replace(/,/g, '')
      .replace(/\(([^)]+)\)/, '-$1')
      .replace(/[^0-9.+-]/g, '');

    if (!cleaned || cleaned === '-' || cleaned === '+') return 0;
    const n = Number(cleaned);
    return Number.isFinite(n) ? n : 0;
  };

  const workbook = periodOverride ? await loadHistoryWorkbook(force) : await loadTotalsWorkbook(force);
  const sheet = findSheetByName(workbook, historySheetName('Receiving1', periodOverride));
  if (!sheet) {
    console.warn('Receiving1 sheet not found');
    return null;
  }

  const data = window.XLSX.utils.sheet_to_json(sheet, { header: 1, raw: false });
  if (!Array.isArray(data) || data.length < 1) {
    console.warn('Receiving1 sheet has insufficient data');
    return null;
  }

  const { year: filterYear, month: filterMonth } = periodOverride || getCurrentInventoryPeriod();
  const monthMatches = (date) => {
    return date instanceof Date && !Number.isNaN(date.getTime()) && date.getFullYear() === filterYear && date.getMonth() === filterMonth;
  };
  const toIsoFromDate = (dt) => dt instanceof Date && !Number.isNaN(dt.getTime())
    ? `${dt.getFullYear()}-${String(dt.getMonth() + 1).padStart(2, '0')}-${String(dt.getDate()).padStart(2, '0')}`
    : '';
  const resolveIsoDate = (dateCell) => {
    const dateText = String(dateCell || '').trim();
    const parsed = parseXlsxDateCell(dateCell);

    if (parsed) {
      return {
        parsed,
        isoDate: toIsoFromDate(parsed),
        dateLabel: `${parsed.getMonth() + 1}/${parsed.getDate()}`
      };
    }

    const maybeDay = Number(dateText);
    if (!Number.isFinite(maybeDay) || maybeDay < 1 || maybeDay > 31) {
      return { parsed: null, isoDate: '', dateLabel: '' };
    }

    return {
      parsed: null,
      isoDate: `${filterYear}-${String(filterMonth + 1).padStart(2, '0')}-${String(maybeDay).padStart(2, '0')}`,
      dateLabel: `Day ${maybeDay}`
    };
  };

  const headerRow = Array.isArray(data[0]) ? data[0] : [];
  const headerNorm = headerRow.map(h => String(h || '').trim().toLowerCase());
  const findCol = (...names) => {
    for (const name of names) {
      const idx = headerNorm.findIndex(h => h === String(name).toLowerCase());
      if (idx >= 0) return idx;
    }
    return -1;
  };

  const dateCol = findCol('date');
  const trucksCol = findCol('total number of trucks received that day', 'trucks received', 'total trucks', 'trucks');
  const weightCol = findCol('total weight received that day', 'weight received', 'total weight', 'weight', 'lbs', 'pounds');

  const resolvedDateCol = dateCol >= 0 ? dateCol : 0;
  const resolvedTrucksCol = trucksCol >= 0 ? trucksCol : 1;
  const resolvedWeightCol = weightCol >= 0 ? weightCol : 2;
  const hasHeaderRow = dateCol >= 0 || trucksCol >= 0 || weightCol >= 0;

  const breakdownByIsoDate = {};
  const receiving2 = findSheetByName(workbook, historySheetName('Receiving2', periodOverride));
  if (receiving2) {
    const data2 = window.XLSX.utils.sheet_to_json(receiving2, { header: 1, raw: false });
    if (Array.isArray(data2) && data2.length > 0) {
      const header2 = Array.isArray(data2[0]) ? data2[0] : [];
      const header2Norm = header2.map(h => String(h || '').trim().toLowerCase());
      const findCol2 = (...names) => {
        for (const name of names) {
          const idx = header2Norm.findIndex(h => h === String(name).toLowerCase());
          if (idx >= 0) return idx;
        }
        return -1;
      };

      const date2Col = findCol2('date');
      const pile2Col = findCol2('pile number', 'pile #', 'pile');
      const lot2Col = findCol2('material lot #', 'material lot#', 'lot #', 'lot');
      const weight2Col = findCol2('total weight received at that pile', 'total weight received', 'weight received', 'total weight', 'weight', 'lbs', 'pounds');
      const hasHeaderRow2 = date2Col >= 0 || pile2Col >= 0 || lot2Col >= 0 || weight2Col >= 0;

      const resolvedDate2Col = date2Col >= 0 ? date2Col : 0;
      const resolvedPile2Col = pile2Col >= 0 ? pile2Col : 1;
      const resolvedLot2Col = lot2Col >= 0 ? lot2Col : 2;
      const resolvedWeight2Col = weight2Col >= 0 ? weight2Col : 3;

      for (let r2 = hasHeaderRow2 ? 1 : 0; r2 < data2.length; r2++) {
        const row2 = data2[r2];
        if (!Array.isArray(row2)) continue;

        const dateCell2 = row2[resolvedDate2Col];
        const pileCell2 = row2[resolvedPile2Col];
        const lotCell2 = row2[resolvedLot2Col];
        const weightCell2 = row2[resolvedWeight2Col];

        if ((dateCell2 == null || String(dateCell2).trim() === '') && (pileCell2 == null || String(pileCell2).trim() === '') && (lotCell2 == null || String(lotCell2).trim() === '') && (weightCell2 == null || String(weightCell2).trim() === '')) {
          continue;
        }

        const dateText2 = String(dateCell2 || '').trim().toLowerCase();
        if (dateText2.includes('total') || dateText2.includes('sum') || dateText2.includes('average') || dateText2.includes('avg')) {
          continue;
        }

        const resolvedDate = resolveIsoDate(dateCell2);
        if (!resolvedDate.isoDate) continue;

        const pile = normalizePileCode(pileCell2);
        const lot = String(lotCell2 || '').trim();
        const weight = toNum(weightCell2);
        if (!pile && !lot && !weight) continue;

        if (!breakdownByIsoDate[resolvedDate.isoDate]) {
          breakdownByIsoDate[resolvedDate.isoDate] = [];
        }

        breakdownByIsoDate[resolvedDate.isoDate].push({
          pile,
          lot,
          material: getMaterialForLotOrPile(lot, pile),
          weight,
          tons: weight / 2000
        });
      }

      Object.keys(breakdownByIsoDate).forEach(iso => {
        breakdownByIsoDate[iso].sort((a, b) => (b.weight - a.weight) || String(a.pile).localeCompare(String(b.pile)));
      });
    }
  }

  const truckDetailsByIsoDate = {};
  const receiving3 = findSheetByName(workbook, historySheetName('Receiving3', periodOverride));
  if (receiving3) {
    const data3 = window.XLSX.utils.sheet_to_json(receiving3, { header: 1, raw: false });
    if (Array.isArray(data3) && data3.length > 0) {
      const header3 = Array.isArray(data3[0]) ? data3[0] : [];
      const header3Norm = header3.map(h => String(h || '').trim().toLowerCase());
      const findCol3 = (...names) => {
        for (const name of names) {
          const idx = header3Norm.findIndex(h => h === String(name).toLowerCase());
          if (idx >= 0) return idx;
        }
        return -1;
      };

      const truckCol3 = findCol3('truck number', 'truck #', 'truck');
      const ticketCol3 = findCol3('ticketid', 'ticket id', 'ticket');
      const pileCol3 = findCol3('pile number the truck dumped at', 'pile number', 'pile #', 'pile');
      const dateCol3 = findCol3('date of material receipt', 'date');
      const lotCol3 = findCol3('material lot #', 'material lot#', 'lot #', 'lot');
      const hasHeaderRow3 = truckCol3 >= 0 || ticketCol3 >= 0 || pileCol3 >= 0 || dateCol3 >= 0 || lotCol3 >= 0;

      const resolvedTruckCol3 = truckCol3 >= 0 ? truckCol3 : 0;
      const resolvedTicketCol3 = ticketCol3 >= 0 ? ticketCol3 : 1;
      const resolvedPileCol3 = pileCol3 >= 0 ? pileCol3 : 2;
      const resolvedDateCol3 = dateCol3 >= 0 ? dateCol3 : 3;
      const resolvedLotCol3 = lotCol3 >= 0 ? lotCol3 : 4;
      const remarksCol3 = findCol3('remarks', 'remark', 'notes', 'note');
      const grossCol3 = findCol3('gross', 'gross weight');
      const tareCol3 = findCol3('tare', 'tare weight');
      const netCol3 = findCol3('net', 'net weight');
      const resolvedRemarksCol3 = remarksCol3 >= 0 ? remarksCol3 : 5;
      const resolvedGrossCol3 = grossCol3 >= 0 ? grossCol3 : 6;
      const resolvedTareCol3 = tareCol3 >= 0 ? tareCol3 : 7;
      const resolvedNetCol3 = netCol3 >= 0 ? netCol3 : 8;

      for (let r3 = hasHeaderRow3 ? 1 : 0; r3 < data3.length; r3++) {
        const row3 = data3[r3];
        if (!Array.isArray(row3)) continue;

        const truckCell3 = row3[resolvedTruckCol3];
        const ticketCell3 = row3[resolvedTicketCol3];
        const pileCell3 = row3[resolvedPileCol3];
        const dateCell3 = row3[resolvedDateCol3];
        const lotCell3 = row3[resolvedLotCol3];

        if ((truckCell3 == null || String(truckCell3).trim() === '') && (ticketCell3 == null || String(ticketCell3).trim() === '') && (pileCell3 == null || String(pileCell3).trim() === '') && (dateCell3 == null || String(dateCell3).trim() === '') && (lotCell3 == null || String(lotCell3).trim() === '')) {
          continue;
        }

        const dateText3 = String(dateCell3 || '').trim().toLowerCase();
        if (dateText3.includes('total') || dateText3.includes('sum') || dateText3.includes('average') || dateText3.includes('avg')) {
          continue;
        }

        const resolvedDate = resolveIsoDate(dateCell3);
        if (!resolvedDate.isoDate) continue;

        const pile = normalizePileCode(pileCell3);
        const lot = String(lotCell3 || '').trim();
        const truckNumber = String(truckCell3 || '').trim();
        const ticketId = String(ticketCell3 || '').trim();
        const remarks = String(row3[resolvedRemarksCol3] || '').trim();
        const gross = String(row3[resolvedGrossCol3] || '').trim();
        const tare = String(row3[resolvedTareCol3] || '').trim();
        const net = String(row3[resolvedNetCol3] || '').trim();

        if (!truckNumber && !ticketId && !pile && !lot) continue;

        if (!truckDetailsByIsoDate[resolvedDate.isoDate]) {
          truckDetailsByIsoDate[resolvedDate.isoDate] = [];
        }

        truckDetailsByIsoDate[resolvedDate.isoDate].push({
          truckNumber,
          ticketId,
          pile,
          material: getMaterialForLotOrPile(lot, pile),
          isStoppedPile: isStoppedPileCode(pile),
          remarks,
          gross,
          tare,
          net
        });
      }

      Object.keys(truckDetailsByIsoDate).forEach(iso => {
        const ticketSortKey = (ticketId) => {
          const raw = String(ticketId || '').trim().toUpperCase();
          const m = raw.match(/^([A-Z]+)(\d+)$/);
          if (m) {
            return { alpha: m[1], num: Number(m[2]), raw };
          }
          return { alpha: '', num: Number.MAX_SAFE_INTEGER, raw };
        };

        truckDetailsByIsoDate[iso].sort((a, b) => {
          const ak = ticketSortKey(a.ticketId);
          const bk = ticketSortKey(b.ticketId);
          return ak.alpha.localeCompare(bk.alpha)
            || (ak.num - bk.num)
            || ak.raw.localeCompare(bk.raw)
            || String(a.truckNumber || '').localeCompare(String(b.truckNumber || ''));
        });
      });
    }
  }

  const rows = [];
  const allRows = [];
  let totalTrucks = 0;
  let totalWeight = 0;

  for (let r = hasHeaderRow ? 1 : 0; r < data.length; r++) {
    const rowData = data[r];
    if (!Array.isArray(rowData)) continue;

    const dateCell = rowData[resolvedDateCol];
    const trucksCell = rowData[resolvedTrucksCol];
    const weightCell = rowData[resolvedWeightCol];

    if ((dateCell == null || String(dateCell).trim() === '') && (trucksCell == null || String(trucksCell).trim() === '') && (weightCell == null || String(weightCell).trim() === '')) {
      continue;
    }

    const dateText = String(dateCell || '').trim();
    const lowerDateText = dateText.toLowerCase();
    if (lowerDateText.includes('total') || lowerDateText.includes('sum') || lowerDateText.includes('average') || lowerDateText.includes('avg')) {
      continue;
    }

    const { parsed, dateLabel, isoDate } = resolveIsoDate(dateCell);
    if (!isoDate) continue;

    const rowObj = {
      dateLabel,
      isoDate,
      trucks: toNum(trucksCell),
      weight: toNum(weightCell),
      breakdown: breakdownByIsoDate[isoDate] || [],
      truckDetails: truckDetailsByIsoDate[isoDate] || []
    };

    allRows.push(rowObj);
    if (monthMatches(parsed) || !parsed) {
      rows.push(rowObj);
    }
  }

  if (rows.length === 0 && allRows.length > 0) {
    rows.push(...allRows);
  }

  rows.sort((a, b) => String(a.isoDate).localeCompare(String(b.isoDate)));

  for (const row of rows) {
    totalTrucks += row.trucks || 0;
    totalWeight += row.weight || 0;
  }

  const payload = {
    rows,
    month: {
      year: filterYear,
      monthIndex: filterMonth,
      rowCount: rows.length
    },
    totals: {
      totalTrucks,
      totalWeight,
      totalTons: totalWeight / 2000
    }
  };

  window.currentReceivingRows = rows;
  if (periodOverride) {
    receivingHistoryCache[`${periodOverride.year}-${periodOverride.month}`] = payload;
  } else {
    receivingCache = { at: now, data: payload };
  }
  return payload;
}

async function fetchRailcarSummary(force = false, periodOverride = null) {
  const now = Date.now();
  if (periodOverride) {
    const key = `${periodOverride.year}-${periodOverride.month}`;
    if (!force && railcarHistoryCache[key]) {
      return railcarHistoryCache[key];
    }
  } else {
    if (!force && railcarCache.data && (now - railcarCache.at) < 120000) {
      return railcarCache.data;
    }
  }

  const workbook = periodOverride ? await loadHistoryWorkbook(force) : await loadTotalsWorkbook(force);
  const sheet = findSheetByName(workbook, historySheetName('Railcars', periodOverride));
  if (!sheet) {
    console.warn('Railcars sheet not found');
    return null;
  }

  const data = window.XLSX.utils.sheet_to_json(sheet, { header: 1, raw: false });
  if (!Array.isArray(data) || data.length < 3) {
    console.warn('Railcars sheet has insufficient data');
    return null;
  }

  const railcars = [];
  let i = 2; // skip 2 header rows

  while (i < data.length) {
    const primary = Array.isArray(data[i]) ? data[i] : [];
    const railcarNum = String(primary[0] || '').trim();

    if (!railcarNum) { i++; continue; }

    const next = (i + 1 < data.length && Array.isArray(data[i + 1])) ? data[i + 1] : [];
    const hasSecondary = !String(next[0] || '').trim(); // secondary row has blank col A

    const statusA = String(primary[11] || '').trim();
    const statusB = hasSecondary ? String(next[11] || '').trim() : '';
    const status = [statusA, statusB].filter(Boolean).join(' ');

    railcars.push({
      railcarNum,
      po:           String(primary[1] || '').trim(),
      supplier:     String(primary[2] || '').trim(),
      lot:          String(primary[3] || '').trim(),
      material:     String(primary[4] || '').trim(),
      shipperGross: String(primary[5] || '').trim(),
      shipperTare:  String(primary[6] || '').trim(),
      shipperNet:   String(primary[7] || '').trim(),
      ourGross:     String(primary[8] || '').trim(),
      ourTare:      String(primary[9] || '').trim(),
      pile:         String(primary[10] || '').trim(),
      status,
      mainSupplier: hasSecondary ? String(next[2] || '').trim() : '',
      consistency:  hasSecondary ? String(next[3] || '').trim() : '',
      regrade:      hasSecondary ? String(next[4] || '').trim() : ''
    });

    i += hasSecondary ? 2 : 1;
  }

  const active = railcars.filter(c => !c.status);
  const released = railcars.filter(c => !!c.status);

  const payload = { railcars, active, released };
  if (periodOverride) {
    const key = `${periodOverride.year}-${periodOverride.month}`;
    railcarHistoryCache[key] = payload;
  } else {
    railcarCache = { at: now, data: payload };
  }
  return payload;
}

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

function renderBucketLoadingPopup(payload) {
  if (!payload) {
    return '<b>Bucket Loading</b><div>No data.</div>';
  }

  const monthLabel = bucketCurrentPeriod
    ? formatMonthYear(bucketCurrentPeriod.year, bucketCurrentPeriod.month)
    : getCurrentInventoryMonthLabel();
  const monthSelectorHtml = buildMonthSelectorHtml('bucket', bucketCurrentPeriod, bucketHistoryMonths);

  const t = payload.totals || {};
  const totalPoundsText = (typeof t.totalPounds === 'number' && isFinite(t.totalPounds)) ? fmtInt(t.totalPounds) : '—';
  const totalTonsText = (typeof t.totalTons === 'number' && isFinite(t.totalTons)) ? fmtTons2(t.totalTons, 2) : '—';
  const avgPoundsText = (typeof t.avgPounds === 'number' && isFinite(t.avgPounds)) ? fmtInt(t.avgPounds) : '—';
  const avgTonsText = (typeof t.avgTons === 'number' && isFinite(t.avgTons)) ? fmtTons2(t.avgTons, 2) : '—';

  const rowsHtml = (payload.rows || []).map(r => `
    <tr class="bucket-loading-main-row" data-date="${esc(r.isoDate)}">
      <td style="padding:2px 6px;white-space:nowrap">${renderDateWithWeekday(r.isoDate, r.dateLabel, { button: true, buttonClass: 'bucket-loading-date-link', dataDate: r.isoDate })}</td>
      <td style="padding:2px 6px;text-align:right">${fmtInt(r.pounds)}</td>
      <td style="padding:2px 6px;text-align:right">${fmtTons2(r.tons, 2)}</td>
      <td style="padding:2px 6px;text-align:right">${fmtInt(r.bucketsLoaded)}</td>
      <td style="padding:2px 6px;text-align:right">${r.heatBreakdown && r.heatBreakdown.length > 0 ? `<button type="button" class="bucket-loading-heat-link" data-date="${esc(r.isoDate)}" style="border:none;background:none;color:#0b57d0;text-decoration:underline;padding:0;cursor:pointer;font:inherit">${fmtInt(r.heatsCompleted)}</button>` : fmtInt(r.heatsCompleted)}</td>
    </tr>
    <tr class="bucket-loading-detail-row" data-date="${esc(r.isoDate)}" style="display:none;background:#fafafa">
      <td colspan="5" style="padding:6px 10px"></td>
    </tr>
    <tr class="bucket-loading-heat-detail-row" data-date="${esc(r.isoDate)}" style="display:none;background:#f0f4ff">
      <td colspan="5" style="padding:6px 10px"></td>
    </tr>
  `).join('');

  const body = `
  <div style="margin-bottom:4px">${monthSelectorHtml}</div>
  <div style="font-weight:700;margin-bottom:6px">Bucket Loading</div>
  <div style="margin-bottom:8px;text-align:center">
    <table style="width:auto;font-size:12px;line-height:1.3;border-collapse:collapse;margin:0 auto;text-align:left">
      <tr><td style="color:#666;padding:2px 6px">Total Consumed</td><td style="text-align:right;padding:2px 6px"><b>${totalPoundsText} lbs</b> <span style="color:#555">(<b>${totalTonsText} tons</b>)</span></td></tr>
      <tr><td style="color:#666;padding:2px 6px">Daily Average</td><td style="text-align:right;padding:2px 6px"><b>${avgPoundsText} lbs</b> <span style="color:#555">(<b>${avgTonsText} tons</b>)</span></td></tr>
    </table>
  </div>
  <div id="bucketLoadingActivity" style="${ACTIVITY_CONTAINER_STYLE}">
    <table style="${ACTIVITY_TABLE_STYLE}">
      <thead>
        <tr style="background:#f2f2f2">
          <th style="text-align:left;padding:2px 6px">Date</th>
          <th style="text-align:right;padding:2px 6px">Pounds</th>
          <th style="text-align:right;padding:2px 6px">Tons</th>
          <th style="text-align:right;padding:2px 6px">Buckets</th>
          <th style="text-align:right;padding:2px 6px">Heats</th>
        </tr>
      </thead>
      <tbody>${rowsHtml}</tbody>
    </table>
  </div>
  <div style="margin-top:6px;display:flex;gap:6px;align-items:center">
    <button type="button" id="bucketLoadingToggle" style="padding:2px 6px;border:1px solid #ddd;border-radius:3px;background:#f8f8f8;cursor:pointer">Show Activity</button>
  </div>
`;

  return `<div style="${POPUP_CONTAINER_STYLE}">${body}</div>`;
}

function renderReceivingPopup(payload) {
  if (!payload) {
    return '<b>Truck Scales</b><div>No data.</div>';
  }

  const monthLabel = receivingCurrentPeriod
    ? formatMonthYear(receivingCurrentPeriod.year, receivingCurrentPeriod.month)
    : getCurrentInventoryMonthLabel();
  const monthSelectorHtml = buildMonthSelectorHtml('receiving', receivingCurrentPeriod, receivingHistoryMonths);

  const t = payload.totals || {};
  const totalTrucksText = (typeof t.totalTrucks === 'number' && isFinite(t.totalTrucks)) ? fmtInt(t.totalTrucks) : '—';
  const totalWeightText = (typeof t.totalWeight === 'number' && isFinite(t.totalWeight)) ? fmtInt(t.totalWeight) : '—';
  const totalTonsText = (typeof t.totalTons === 'number' && isFinite(t.totalTons)) ? fmtTons2(t.totalTons, 2) : '—';
  const weekdayCount = (payload.rows || []).filter(r => {
    if (!r.isoDate) return false;
    const [yr, mo, dy] = r.isoDate.split('-').map(Number);
    const dow = new Date(yr, mo - 1, dy).getDay();
    return dow >= 1 && dow <= 5; // Monday–Friday only
  }).length;
  const dailyAvgWeight = weekdayCount > 0 ? (t.totalWeight / weekdayCount) : 0;
  const dailyAvgWeightText = Number.isFinite(dailyAvgWeight) ? fmtInt(dailyAvgWeight) : '—';
  const dailyAvgTonsText = Number.isFinite(dailyAvgWeight) ? fmtTons2(dailyAvgWeight / 2000, 2) : '—';

  const rowsHtml = (payload.rows || []).map(r => `
    <tr class="receiving-main-row" data-date="${esc(r.isoDate)}">
      <td style="padding:2px 6px;white-space:nowrap">${r.breakdown && r.breakdown.length > 0 ? renderDateWithWeekday(r.isoDate, r.dateLabel, { button: true, buttonClass: 'receiving-date-link', dataDate: r.isoDate }) : renderDateWithWeekday(r.isoDate, r.dateLabel)}</td>
      <td style="padding:2px 6px;text-align:right">${fmtInt(r.weight)}</td>
      <td style="padding:2px 6px;text-align:right">${fmtTons2((r.weight || 0) / 2000, 2)}</td>
      <td style="padding:2px 6px;text-align:right">${r.truckDetails && r.truckDetails.length > 0 ? `<button type="button" class="receiving-trucks-link" data-date="${esc(r.isoDate)}" style="border:none;background:none;color:#0b57d0;text-decoration:underline;padding:0;cursor:pointer;font:inherit">${fmtInt(r.trucks)}</button>` : fmtInt(r.trucks)}</td>
    </tr>
    <tr class="receiving-detail-row" data-date="${esc(r.isoDate)}" style="display:none;background:#fafafa">
      <td colspan="4" style="padding:6px 10px"></td>
    </tr>
    <tr class="receiving-truck-detail-row" data-date="${esc(r.isoDate)}" style="display:none;background:#f3f7ff">
      <td colspan="4" style="padding:6px 10px"></td>
    </tr>
  `).join('');

  const body = `
  <div style="margin-bottom:4px">${monthSelectorHtml}</div>
  <div style="font-weight:700;margin-bottom:6px">Truck Scales</div>
  <div style="margin-bottom:8px;text-align:center">
    <table style="width:auto;font-size:12px;line-height:1.3;border-collapse:collapse;margin:0 auto;text-align:left">
      <tr><td style="color:#666;padding:2px 6px">Total Trucks Received</td><td style="text-align:right;padding:2px 6px"><b>${totalTrucksText}</b></td></tr>
      <tr><td style="color:#666;padding:2px 6px">Total Material Received</td><td style="text-align:right;padding:2px 6px"><b>${totalWeightText} lbs</b> <span style="color:#555">(<b>${totalTonsText} tons</b>)</span></td></tr>
      <tr><td style="color:#666;padding:2px 6px">Daily Average Material Received</td><td style="text-align:right;padding:2px 6px"><b>${dailyAvgWeightText} lbs</b> <span style="color:#555">(<b>${dailyAvgTonsText} tons</b>)</span></td></tr>
    </table>
  </div>
  <div id="receivingActivity" style="${ACTIVITY_CONTAINER_STYLE}">
    <table style="${ACTIVITY_TABLE_STYLE}">
      <thead>
        <tr style="background:#f2f2f2">
          <th style="text-align:left;padding:2px 6px">Date</th>
          <th style="text-align:right;padding:2px 6px">Weight (lbs)</th>
          <th style="text-align:right;padding:2px 6px">Tons</th>
          <th style="text-align:right;padding:2px 6px">Trucks</th>
        </tr>
      </thead>
      <tbody>${rowsHtml}</tbody>
    </table>
  </div>
  <div style="margin-top:6px;display:flex;gap:6px;align-items:center">
    <button type="button" id="receivingToggle" style="padding:2px 6px;border:1px solid #ddd;border-radius:3px;background:#f8f8f8;cursor:pointer">Show Activity</button>
  </div>
`;

  return `<div style="${POPUP_CONTAINER_STYLE}">${body}</div>`;
}

function showRailcarPhoto(railcarNum, monthFolder) {
  if (!railcarPhotoOverlay) {
    const overlay = document.createElement('div');
    overlay.id = 'railcar-photo-overlay';
    overlay.style.cssText = 'position:fixed;top:0;left:0;width:100%;height:100%;background:rgba(0,0,0,0.82);z-index:9999;display:flex;flex-direction:column;align-items:center;justify-content:center;cursor:pointer';
    overlay.innerHTML = '<img id="railcar-photo-img" style="max-width:90vw;max-height:80vh;object-fit:contain;box-shadow:0 4px 32px rgba(0,0,0,0.7);border-radius:4px">'
      + '<div id="railcar-photo-caption" style="color:#fff;margin-top:10px;font-size:14px;font-family:sans-serif;text-align:center"></div>'
      + '<div style="color:#aaa;font-size:12px;margin-top:4px;font-family:sans-serif">Click anywhere or press ESC to close</div>';
    overlay.addEventListener('click', () => { overlay.style.display = 'none'; });
    document.addEventListener('keydown', e => { if (e.key === 'Escape') overlay.style.display = 'none'; });
    document.body.appendChild(overlay);
    railcarPhotoOverlay = overlay;
  }

  const img = railcarPhotoOverlay.querySelector('#railcar-photo-img');
  const caption = railcarPhotoOverlay.querySelector('#railcar-photo-caption');
  img.src = '';
  caption.textContent = 'Loading…';
  railcarPhotoOverlay.style.display = 'flex';

  function tryNext(i) {
    const exts = ['png', 'jpg', 'jpeg'];
    if (i >= exts.length) {
      img.src = '';
      caption.textContent = 'No photo found for ' + railcarNum;
      return;
    }
    const url = 'RailcarPhotos/' + monthFolder + '/' + railcarNum + '.' + exts[i];
    const test = new Image();
    test.onload = () => { img.src = url; caption.textContent = railcarNum + ' — ' + monthFolder; };
    test.onerror = () => tryNext(i + 1);
    test.src = url;
  }
  tryNext(0);
}

function renderRailcarPopup(payload) {
  if (!payload) {
    return '<b>Railroad</b><div>No data.</div>';
  }

  const active = payload.active || [];
  const released = payload.released || [];
  const total = Array.isArray(payload.railcars) ? payload.railcars.length : 0;

  const parseWeight = s => {
    if (!s) return NaN;
    const n = Number(String(s).replace(/,/g, '').trim());
    return Number.isFinite(n) ? n : NaN;
  };
  const fmtW = s => {
    const n = parseWeight(s);
    return Number.isFinite(n) ? fmtInt(n) : (s || '—');
  };
  const computeOurNet = car => {
    const g = parseWeight(car.ourGross);
    const t = parseWeight(car.ourTare);
    if (Number.isFinite(g) && Number.isFinite(t)) return fmtInt(g - t);
    return '—';
  };

  const fmtSupplier = car => {
    const line1 = esc(car.supplier || '—');
    if (!car.mainSupplier) return line1;
    return line1 + '<br><span style="font-size:11px;color:#777">' + esc(car.mainSupplier) + '</span>';
  };

  const fmtMaterial = car => {
    const line1 = esc(car.material || '—');
    const extras = [car.regrade, car.consistency].filter(Boolean).map(esc).join(' · ');
    if (!extras) return line1;
    return line1 + '<br><span style="font-size:11px;color:#777">' + extras + '</span>';
  };

  const monthFolder = getCurrentInventoryMonthLabel();
  const monthLabel = railcarCurrentPeriod
    ? formatMonthYear(railcarCurrentPeriod.year, railcarCurrentPeriod.month)
    : getCurrentInventoryMonthLabel();
  const monthSelectorHtml = buildMonthSelectorHtml('railcar', railcarCurrentPeriod, railcarHistoryMonths);

  const activeHtml = active.length > 0
    ? '<table style="width:100%;font-size:12px;border-collapse:collapse">'
      + '<thead><tr style="background:#fff3e0">'
      + '<th style="text-align:left;padding:2px 6px">Car #</th>'
      + '<th style="text-align:left;padding:2px 6px">PO #</th>'
      + '<th style="text-align:left;padding:2px 6px">Supplier</th>'
      + '<th style="text-align:left;padding:2px 6px">Material</th>'
      + '<th style="text-align:right;padding:2px 6px">Shpr Gross</th>'
      + '<th style="text-align:right;padding:2px 6px">Shpr Tare</th>'
      + '<th style="text-align:right;padding:2px 6px">Shpr Net</th>'
      + '<th style="text-align:right;padding:2px 6px">Scale Gross</th>'
      + '</tr></thead><tbody>'
      + active.map(car => {
          const hasScale = !!car.ourGross;
          const rowStyle = 'border-top:1px solid #f5f5f5;' + (hasScale ? 'background:#e8f5e9;' : '');
          const gOurs = parseWeight(car.ourGross);
          const gShpr = parseWeight(car.shipperGross);
          const grossDiff = (Number.isFinite(gOurs) && Number.isFinite(gShpr)) ? (gOurs - gShpr) : null;
          const grossTitle = grossDiff !== null
            ? fmtInt(Math.abs(grossDiff)) + ' lbs ' + (grossDiff > 0 ? "Higher than shipper's" : grossDiff < 0 ? "Lower than shipper's" : 'Exact match')
            : '';
          const scaleGrossCell = hasScale
            ? (grossTitle ? '<span style="text-decoration:underline dotted">' + fmtW(car.ourGross) + '</span>' : fmtW(car.ourGross))
            : '—';
          return '<tr style="' + rowStyle + '">'
            + '<td style="padding:2px 6px;font-weight:600"><button type="button" class="railcar-photo-btn" data-car="' + esc(car.railcarNum) + '" data-month="' + esc(monthFolder) + '" style="background:none;border:none;padding:0;cursor:pointer;font-weight:600;color:#0b57d0;text-decoration:underline;font-size:inherit">' + esc(car.railcarNum) + '</button></td>'
            + '<td style="padding:2px 6px;color:#555">' + esc(car.po || '—') + '</td>'
            + '<td style="padding:2px 6px">' + fmtSupplier(car) + '</td>'
            + '<td style="padding:2px 6px">' + fmtMaterial(car) + '</td>'
            + '<td style="padding:2px 6px;text-align:right">' + fmtW(car.shipperGross) + '</td>'
            + '<td style="padding:2px 6px;text-align:right">' + fmtW(car.shipperTare) + '</td>'
            + '<td style="padding:2px 6px;text-align:right">' + fmtW(car.shipperNet) + '</td>'
            + '<td style="padding:2px 6px;text-align:right"' + (grossTitle ? ' data-offset="' + grossTitle + '"' : '') + '>' + scaleGrossCell + '</td>'
            + '</tr>';
        }).join('')
      + '</tbody></table>'
    : '<div style="font-size:12px;color:#888;padding:6px 0">No active railcars.</div>';

  const releasedHtml = released.length > 0
    ? '<table style="width:100%;font-size:12px;border-collapse:collapse">'
      + '<thead><tr style="background:#e8f5e9">'
      + '<th style="text-align:left;padding:2px 6px">Car #</th>'
      + '<th style="text-align:left;padding:2px 6px">PO #</th>'
      + '<th style="text-align:left;padding:2px 6px">Supplier</th>'
      + '<th style="text-align:left;padding:2px 6px">Material</th>'
      + '<th style="text-align:right;padding:2px 6px">Scale Gross</th>'
      + '<th style="text-align:right;padding:2px 6px">Scale Tare</th>'
      + '<th style="text-align:right;padding:2px 6px">Our Net (lbs)</th>'
      + '<th style="text-align:left;padding:2px 6px">Pile</th>'
      + '<th style="text-align:left;padding:2px 6px">Status</th>'
      + '</tr></thead><tbody>'
      + released.map(car => {
          const gOurs = parseWeight(car.ourGross);
          const gShpr = parseWeight(car.shipperGross);
          const grossDiff = (Number.isFinite(gOurs) && Number.isFinite(gShpr)) ? (gOurs - gShpr) : null;
          const grossTitle = grossDiff !== null
            ? fmtInt(Math.abs(grossDiff)) + ' lbs ' + (grossDiff > 0 ? "Higher than shipper's" : grossDiff < 0 ? "Lower than shipper's" : 'Exact match')
            : '';
          const scaleGrossCell = car.ourGross
            ? (grossTitle ? '<span style="text-decoration:underline dotted">' + fmtW(car.ourGross) + '</span>' : fmtW(car.ourGross))
            : '—';
          const ourNetNum = parseWeight(car.ourGross) - parseWeight(car.ourTare);
          const shprNetNum = parseWeight(car.shipperNet);
          const netDiff = (Number.isFinite(ourNetNum) && Number.isFinite(shprNetNum)) ? (ourNetNum - shprNetNum) : null;
          const netTitle = netDiff !== null
            ? fmtInt(Math.abs(netDiff)) + ' lbs ' + (netDiff > 0 ? "Higher than shipper's" : netDiff < 0 ? "Lower than shipper's" : 'Exact match')
            : '';
          const ourNetCell = netTitle
            ? '<span style="text-decoration:underline dotted">' + computeOurNet(car) + '</span>'
            : computeOurNet(car);
          return '<tr style="border-top:1px solid #f5f5f5">'
            + '<td style="padding:2px 6px;font-weight:600"><button type="button" class="railcar-photo-btn" data-car="' + esc(car.railcarNum) + '" data-month="' + esc(monthFolder) + '" style="background:none;border:none;padding:0;cursor:pointer;font-weight:600;color:#0b57d0;text-decoration:underline;font-size:inherit">' + esc(car.railcarNum) + '</button></td>'
            + '<td style="padding:2px 6px;color:#555">' + esc(car.po || '—') + '</td>'
            + '<td style="padding:2px 6px">' + fmtSupplier(car) + '</td>'
            + '<td style="padding:2px 6px">' + fmtMaterial(car) + '</td>'
            + '<td style="padding:2px 6px;text-align:right"' + (grossTitle ? ' data-offset="' + grossTitle + '"' : '') + '>' + scaleGrossCell + '</td>'
            + '<td style="padding:2px 6px;text-align:right">' + (car.ourTare ? fmtW(car.ourTare) : '—') + '</td>'
            + '<td style="padding:2px 6px;text-align:right"' + (netTitle ? ' data-offset="' + netTitle + '"' : '') + '>' + ourNetCell + '</td>'
            + '<td style="padding:2px 6px">' + esc(car.pile || '—') + '</td>'
            + '<td style="padding:2px 6px">' + esc(car.status || '—') + '</td>'
            + '</tr>';
        }).join('')
      + '</tbody></table>'
    : '<div style="font-size:12px;color:#888;padding:6px 0">No released railcars.</div>';

  const railNetTotal = released.reduce((sum, car) => {
    const g = parseWeight(car.ourGross);
    const t = parseWeight(car.ourTare);
    return (Number.isFinite(g) && Number.isFinite(t)) ? sum + (g - t) : sum;
  }, 0);
  const railNetTotalText = railNetTotal > 0
    ? `<b>${fmtInt(railNetTotal)} lbs</b> <span style="color:#555">(<b>${fmtTons2(railNetTotal / 2000, 2)} tons</b>)</span>`
    : '—';

  const body = `
    <div style="margin-bottom:4px">${monthSelectorHtml}</div>
    <div style="font-weight:700;margin-bottom:6px">Railroad</div>
    <div style="margin-bottom:8px;text-align:center">
      <table style="width:auto;font-size:12px;line-height:1.3;border-collapse:collapse;margin:0 auto;text-align:left">
        <tr><td style="color:#666;padding:2px 6px">Total Railcars</td><td style="text-align:right;padding:2px 6px"><b>${total}</b></td></tr>
        <tr><td style="color:#e65100;padding:2px 6px">&#9679; Active (In Yard)</td><td style="text-align:right;padding:2px 6px"><b>${active.length}</b></td></tr>
        <tr><td style="color:#2e7d32;padding:2px 6px">&#9679; Released</td><td style="text-align:right;padding:2px 6px"><b>${released.length}</b></td></tr>
        <tr><td style="color:#666;padding:2px 6px;padding-top:6px;border-top:1px solid #eee">Total Weight Received</td><td style="text-align:right;padding:2px 6px;padding-top:6px;border-top:1px solid #eee">${railNetTotalText}</td></tr>
      </table>
    </div>
    <div style="display:flex;gap:6px;align-items:center;flex-wrap:wrap">
      <button type="button" id="railcarOnsiteToggle" style="padding:2px 6px;border:1px solid #ddd;border-radius:3px;background:#f8f8f8;cursor:pointer">Railcars Onsite (${active.length})</button>
      <button type="button" id="railcarReleasedToggle" style="padding:2px 6px;border:1px solid #ddd;border-radius:3px;background:#f8f8f8;cursor:pointer">Released Railcars (${released.length})</button>
    </div>
    <div id="railcarOnsiteSection" style="display:none;margin-top:6px">
      <div style="font-size:11px;font-weight:700;color:#e65100;margin-bottom:4px;text-transform:uppercase;letter-spacing:0.06em">Railcars Onsite (${active.length})</div>
      <div style="max-height:220px;overflow-y:auto;border:1px solid #ffe0b2;border-radius:3px">
        ${activeHtml}
      </div>
    </div>
    <div id="railcarReleasedSection" style="display:none;margin-top:6px">
      <div style="font-size:11px;font-weight:700;color:#2e7d32;margin-bottom:4px;text-transform:uppercase;letter-spacing:0.06em">Released Railcars (${released.length})</div>
      <div style="max-height:260px;overflow-y:auto;border:1px solid #c8e6c9;border-radius:3px">
        ${releasedHtml}
      </div>
    </div>
  `;

  return `<div style="min-width:640px;max-width:100%;width:auto">${body}</div>`;
}

function wireBucketLoadingPopupEvents(container, marker) {
  const monthBtn = container.querySelector('#bucketMonthBtn');
  if (monthBtn) {
    monthBtn.addEventListener('click', async (e) => {
      e.stopPropagation(); e.preventDefault();
      monthBtn.textContent = 'Loading…';
      monthBtn.disabled = true;
      if (!bucketHistoryMonths) bucketHistoryMonths = await discoverHistoryMonths('Consumption1');
      const payload = await fetchBucketLoadingConsumption(false, bucketCurrentPeriod);
      marker.setPopupContent(unescapeAngles(renderBucketLoadingPopup(payload)));
      setTimeout(() => { const el = marker.getPopup()?.getElement(); if (el) wireBucketLoadingPopupEvents(el, marker); }, 0);
    });
  }
  const calToggle = container.querySelector('#bucketMonthCalendarToggle');
  if (calToggle) {
    calToggle.addEventListener('click', (e) => {
      e.preventDefault(); e.stopPropagation();
      const cal = container.querySelector('#bucketMonthCalendar');
      if (cal) cal.style.display = cal.style.display === 'none' ? '' : 'none';
    });
  }
  container.querySelectorAll('.bucketMonthCell').forEach(cell => {
    cell.addEventListener('click', async (e) => {
      e.preventDefault(); e.stopPropagation();
      const val = cell.getAttribute('data-value');
      bucketCurrentPeriod = val === 'current' ? null : { year: +val.split('-')[0], month: +val.split('-')[1] };
      const payload = await fetchBucketLoadingConsumption(false, bucketCurrentPeriod);
      marker.setPopupContent(unescapeAngles(renderBucketLoadingPopup(payload)));
      setTimeout(() => { const el = marker.getPopup()?.getElement(); if (el) wireBucketLoadingPopupEvents(el, marker); }, 0);
    });
  });
  const bucketResetBtn = container.querySelector('#bucketMonthResetBtn');
  if (bucketResetBtn) {
    bucketResetBtn.addEventListener('click', async (e) => {
      e.preventDefault(); e.stopPropagation();
      bucketCurrentPeriod = null;
      const payload = await fetchBucketLoadingConsumption(false, null);
      marker.setPopupContent(unescapeAngles(renderBucketLoadingPopup(payload)));
      setTimeout(() => { const el = marker.getPopup()?.getElement(); if (el) wireBucketLoadingPopupEvents(el, marker); }, 0);
    });
  }

  const toggle = container.querySelector('#bucketLoadingToggle');
  const block = container.querySelector('#bucketLoadingActivity');

  function setBucketLinkState(link, isActive) {
    if (!link) return;
    link.style.color = isActive ? '#c62828' : '#0b57d0';
    link.style.fontWeight = isActive ? '700' : '400';
  }

  function clearBucketLinkStates() {
    container.querySelectorAll('.bucket-loading-date-link, .bucket-loading-heat-link').forEach(link => {
      setBucketLinkState(link, false);
    });
  }

  function setActiveBucketLink(selector, dateKey) {
    clearBucketLinkStates();
    if (!dateKey) return;
    const activeLink = container.querySelector(`${selector}[data-date="${dateKey}"]`);
    if (activeLink) setBucketLinkState(activeLink, true);
  }

  if (block) block.style.display = 'none';

  if (toggle && block) {
    toggle.addEventListener('click', (e) => {
      e.preventDefault();
      e.stopPropagation();
      const isHidden = block.style.display === 'none';
      block.style.display = isHidden ? '' : 'none';
      toggle.textContent = isHidden ? 'Hide Activity' : 'Show Activity';
    });
    block.addEventListener('click', e => e.stopPropagation());
  }

  // Closes every expanded detail row of either type, optionally skipping one element.
  function closeAllDetailRows(except) {
    container.querySelectorAll('.bucket-loading-detail-row, .bucket-loading-heat-detail-row').forEach(dr => {
      if (dr !== except) dr.style.display = 'none';
    });
    if (!except) clearBucketLinkStates();
  }

  const dateLinks = container.querySelectorAll('.bucket-loading-date-link');
  if (dateLinks && dateLinks.length > 0) {
    dateLinks.forEach(link => {
      link.addEventListener('click', (e) => {
        e.preventDefault();
        e.stopPropagation();

        const dateKey = String(link.getAttribute('data-date') || '').trim();
        if (!dateKey) return;

        const rows = Array.isArray(window.currentBucketLoadingRows) ? window.currentBucketLoadingRows : [];
        const dayRow = rows.find(r => String(r.isoDate || '') === dateKey);
        if (!dayRow) return;

        const detailRow = Array.from(container.querySelectorAll('.bucket-loading-detail-row'))
          .find(dr => dr.getAttribute('data-date') === dateKey);
        if (!detailRow) return;

        const currentlyVisible = detailRow.style.display !== 'none';
        // Close everything (including this row if toggling off)
        closeAllDetailRows(null);
        if (currentlyVisible) return;

        detailRow.style.display = '';
        setActiveBucketLink('.bucket-loading-date-link', dateKey);

        const detailCell = detailRow.querySelector('td');
        if (!detailCell) return;

        const breakdownRows = Array.isArray(dayRow.breakdown) ? dayRow.breakdown : [];
        const breakdownTotal = breakdownRows.reduce((sum, item) => sum + (Number(item.pounds) || 0), 0);
        const breakdownTotalTons = breakdownRows.reduce((sum, item) => sum + (Number(item.tons) || 0), 0);
        const breakdownHtml = breakdownRows.length > 0
          ? breakdownRows.map(item => {
              const lots = Array.isArray(item.lots) ? item.lots : [];
              const lotsText = lots.length > 0
                ? lots.map(x => esc(x.material || '') || '—').join('<br>')
                : '—';
              return `
              <tr>
                <td style="padding:4px 8px;border-bottom:1px solid #eee;white-space:nowrap">${esc(item.pile || '')}</td>
                <td style="padding:4px 8px;border-bottom:1px solid #eee;white-space:nowrap">${lotsText}</td>
                <td style="padding:4px 8px;border-bottom:1px solid #eee;text-align:right">${fmtInt(item.pounds)}</td>
                <td style="padding:4px 8px;border-bottom:1px solid #eee;text-align:right">${fmtTons2(item.tons, 2)}</td>
              </tr>`;
            }).join('')
          : '<tr><td colspan="4" style="padding:8px;color:#666">No pile breakdown data for this day.</td></tr>';

        const totalsHtml = `
          <tr style="background:#f8f8f8">
            <td style="padding:5px 8px;border-top:1px solid #ddd"><b>Daily Total</b></td>
            <td style="padding:5px 8px;border-top:1px solid #ddd">&nbsp;</td>
            <td style="padding:5px 8px;border-top:1px solid #ddd;text-align:right"><b>${fmtInt(breakdownTotal)}</b></td>
            <td style="padding:5px 8px;border-top:1px solid #ddd;text-align:right"><b>${fmtTons2(breakdownTotalTons, 2)}</b></td>
          </tr>`;

        detailCell.innerHTML = `
          <div style="font-size:12px;color:#444;margin-bottom:4px"><b>${esc(dayRow.dateLabel)}</b> pile usage breakdown</div>
          <table style="width:100%;border-collapse:collapse;font-size:12px;background:#fff">
            <thead>
              <tr style="background:#f4f6f8">
                <th style="padding:4px 8px;text-align:left;border-bottom:1px solid #ddd">Pile Utilized</th>
                <th style="padding:4px 8px;text-align:left;border-bottom:1px solid #ddd">Material</th>
                <th style="padding:4px 8px;text-align:right;border-bottom:1px solid #ddd">Pounds</th>
                <th style="padding:4px 8px;text-align:right;border-bottom:1px solid #ddd">Tons</th>
              </tr>
            </thead>
            <tbody>${breakdownHtml}${totalsHtml}</tbody>
          </table>`;
      });
    });
  }

  const heatLinks = container.querySelectorAll('.bucket-loading-heat-link');
  if (heatLinks && heatLinks.length > 0) {
    heatLinks.forEach(link => {
      link.addEventListener('click', (e) => {
        e.preventDefault();
        e.stopPropagation();

        const dateKey = String(link.getAttribute('data-date') || '').trim();
        if (!dateKey) return;

        const rowsData = Array.isArray(window.currentBucketLoadingRows) ? window.currentBucketLoadingRows : [];
        const dayRow = rowsData.find(r => String(r.isoDate || '') === dateKey);
        if (!dayRow) return;

        const heatDetailRow = Array.from(container.querySelectorAll('.bucket-loading-heat-detail-row'))
          .find(dr => dr.getAttribute('data-date') === dateKey);
        if (!heatDetailRow) return;

        const currentlyVisible = heatDetailRow.style.display !== 'none';
        // Close everything (including this row if toggling off)
        closeAllDetailRows(null);
        if (currentlyVisible) return;

        heatDetailRow.style.display = '';
        setActiveBucketLink('.bucket-loading-heat-link', dateKey);

        const detailCell = heatDetailRow.querySelector('td');
        if (!detailCell) return;

        const heats = Array.isArray(dayRow.heatBreakdown) ? dayRow.heatBreakdown : [];
        if (heats.length === 0) {
          detailCell.innerHTML = '<div style="padding:8px;color:#666;font-size:12px">No heat detail data for this day.</div>';
          return;
        }

        // Build accordion: one header row per heat, collapsed material rows below each.
        const accordionRowsHtml = heats.map((heat, hIdx) => {
          const mats = Array.isArray(heat.materials) ? heat.materials : [];
          const heatTotalLbs  = mats.reduce((s, m) => s + (Number(m.pounds) || 0), 0);
          const heatTotalTons = mats.reduce((s, m) => s + (Number(m.tons)   || 0), 0);

          const matsHtml = mats.length > 0
            ? mats.map(m => `
                <tr class="heat-mat-row" data-heat-idx="${hIdx}" style="display:none;background:#f7f8ff">
                  <td style="padding:3px 8px 3px 24px;border-bottom:1px solid #e8edf8;white-space:nowrap">${esc(m.pile || '—')}</td>
                  <td style="padding:3px 8px;border-bottom:1px solid #e8edf8;white-space:normal;overflow-wrap:anywhere;line-height:1.2">
                    <div>${esc(m.material || '—')}</div>
                  </td>
                  <td style="padding:3px 8px;border-bottom:1px solid #e8edf8;text-align:right;white-space:nowrap">${fmtInt(m.pounds)}</td>
                  <td style="padding:3px 8px;border-bottom:1px solid #e8edf8;text-align:right;white-space:nowrap">${fmtTons2(m.tons, 3)}</td>
                </tr>`).join('')
            : `<tr class="heat-mat-row" data-heat-idx="${hIdx}" style="display:none;background:#f7f8ff">
                 <td colspan="4" style="padding:4px 8px 4px 24px;color:#888;font-style:italic">No material records</td>
               </tr>`;

          const chevron = mats.length > 0
            ? `<span class="heat-chevron" style="font-size:10px;margin-right:4px;display:inline-block;transition:transform 0.15s">▶</span>`
            : `<span style="font-size:10px;margin-right:4px;color:#bbb">—</span>`;

          return `
            <tr class="heat-accordion-header" data-heat-idx="${hIdx}"
                style="cursor:${mats.length > 0 ? 'pointer' : 'default'};border-bottom:1px solid #c8d4f0">
              <td style="padding:5px 8px;white-space:nowrap;font-weight:600">${chevron}${esc(heat.heatNumber)}</td>
              <td style="padding:5px 8px;white-space:nowrap">${esc(heat.grade)}</td>
              <td style="padding:5px 8px;text-align:right">${fmtInt(heatTotalLbs)}</td>
              <td style="padding:5px 8px;text-align:right">${fmtTons2(heatTotalTons, 3)}</td>
            </tr>
            <tr class="heat-mat-subheader heat-mat-row" data-heat-idx="${hIdx}" style="display:none;background:#dce4f7;font-size:11px">
              <th style="padding:2px 8px 2px 24px;text-align:left;font-weight:600;border-bottom:1px solid #c8d4f0">Pile</th>
              <th style="padding:2px 8px;text-align:left;font-weight:600;border-bottom:1px solid #c8d4f0">Material</th>
              <th style="padding:2px 8px;text-align:right;font-weight:600;border-bottom:1px solid #c8d4f0">Pounds</th>
              <th style="padding:2px 8px;text-align:right;font-weight:600;border-bottom:1px solid #c8d4f0">Tons</th>
            </tr>
            ${matsHtml}`;
        }).join('');

        detailCell.innerHTML = `
          <div style="font-size:12px;color:#444;margin-bottom:4px">
            <b>${fmtInt(heats.length)} heat${heats.length !== 1 ? 's' : ''} completed on ${esc(dayRow.dateLabel)}</b>
            <span style="color:#888;font-weight:normal"> -- Click a row to expand</span>
          </div>
          <div style="width:100%;max-width:100%;overflow-x:auto">
          <table style="width:100%;table-layout:fixed;border-collapse:collapse;font-size:12px;background:#fff">
            <thead>
              <tr style="background:#e8edf8">
                <th style="padding:4px 8px;text-align:left;border-bottom:1px solid #c8d4f0">Heat #</th>
                <th style="padding:4px 8px;text-align:left;border-bottom:1px solid #c8d4f0">Grade</th>
                <th style="padding:4px 8px;text-align:right;border-bottom:1px solid #c8d4f0">Total Lbs</th>
                <th style="padding:4px 8px;text-align:right;border-bottom:1px solid #c8d4f0">Total Tons</th>
              </tr>
            </thead>
            <tbody>${accordionRowsHtml}</tbody>
          </table>
          </div>`;

        // Wire accordion toggle for each heat header row
        detailCell.querySelectorAll('.heat-accordion-header').forEach(headerRow => {
          const hIdx = headerRow.getAttribute('data-heat-idx');
          const heat = heats[Number(hIdx)];
          if (!heat || !Array.isArray(heat.materials) || heat.materials.length === 0) return;

          headerRow.addEventListener('click', () => {
            const matRows = detailCell.querySelectorAll(`.heat-mat-row[data-heat-idx="${hIdx}"]`);
            const chevron = headerRow.querySelector('.heat-chevron');
            const isOpen = matRows.length > 0 && matRows[0].style.display !== 'none';

            // Collapse all other open heats first
            detailCell.querySelectorAll('.heat-accordion-header').forEach(otherHeader => {
              const oIdx = otherHeader.getAttribute('data-heat-idx');
              if (oIdx === hIdx) return;
              detailCell.querySelectorAll(`.heat-mat-row[data-heat-idx="${oIdx}"]`)
                .forEach(r => { r.style.display = 'none'; });
              const otherChevron = otherHeader.querySelector('.heat-chevron');
              if (otherChevron) otherChevron.style.transform = '';
            });

            matRows.forEach(r => { r.style.display = isOpen ? 'none' : ''; });
            if (chevron) chevron.style.transform = isOpen ? '' : 'rotate(90deg)';
          });
        });
      });
    });
  }
}

function wireReceivingPopupEvents(container, marker) {
  const monthBtn = container.querySelector('#receivingMonthBtn');
  if (monthBtn) {
    monthBtn.addEventListener('click', async (e) => {
      e.stopPropagation(); e.preventDefault();
      monthBtn.textContent = 'Loading…';
      monthBtn.disabled = true;
      if (!receivingHistoryMonths) receivingHistoryMonths = await discoverHistoryMonths('Receiving1');
      const payload = await fetchReceivingSummary(false, receivingCurrentPeriod);
      marker.setPopupContent(unescapeAngles(renderReceivingPopup(payload)));
      setTimeout(() => { const el = marker.getPopup()?.getElement(); if (el) wireReceivingPopupEvents(el, marker); }, 0);
    });
  }
  const receivingCalToggle = container.querySelector('#receivingMonthCalendarToggle');
  if (receivingCalToggle) {
    receivingCalToggle.addEventListener('click', (e) => {
      e.preventDefault(); e.stopPropagation();
      const cal = container.querySelector('#receivingMonthCalendar');
      if (cal) cal.style.display = cal.style.display === 'none' ? '' : 'none';
    });
  }
  container.querySelectorAll('.receivingMonthCell').forEach(cell => {
    cell.addEventListener('click', async (e) => {
      e.preventDefault(); e.stopPropagation();
      const val = cell.getAttribute('data-value');
      receivingCurrentPeriod = val === 'current' ? null : { year: +val.split('-')[0], month: +val.split('-')[1] };
      const payload = await fetchReceivingSummary(false, receivingCurrentPeriod);
      marker.setPopupContent(unescapeAngles(renderReceivingPopup(payload)));
      setTimeout(() => { const el = marker.getPopup()?.getElement(); if (el) wireReceivingPopupEvents(el, marker); }, 0);
    });
  });
  const receivingResetBtn = container.querySelector('#receivingMonthResetBtn');
  if (receivingResetBtn) {
    receivingResetBtn.addEventListener('click', async (e) => {
      e.preventDefault(); e.stopPropagation();
      receivingCurrentPeriod = null;
      const payload = await fetchReceivingSummary(false, null);
      marker.setPopupContent(unescapeAngles(renderReceivingPopup(payload)));
      setTimeout(() => { const el = marker.getPopup()?.getElement(); if (el) wireReceivingPopupEvents(el, marker); }, 0);
    });
  }

  const toggle = container.querySelector('#receivingToggle');
  const block = container.querySelector('#receivingActivity');

  function setReceivingLinkState(link, isActive) {
    if (!link) return;
    link.style.color = isActive ? '#c62828' : '#0b57d0';
    link.style.fontWeight = isActive ? '700' : '400';
  }

  function clearReceivingLinkStates() {
    container.querySelectorAll('.receiving-date-link, .receiving-trucks-link').forEach(link => {
      setReceivingLinkState(link, false);
    });
  }

  function closeAllDetailRows(except) {
    container.querySelectorAll('.receiving-detail-row, .receiving-truck-detail-row').forEach(row => {
      if (row !== except) row.style.display = 'none';
    });
    if (!except) clearReceivingLinkStates();
  }

  if (toggle && block) {
    toggle.addEventListener('click', (e) => {
      e.preventDefault();
      e.stopPropagation();
      const isHidden = block.style.display === 'none';
      block.style.display = isHidden ? 'block' : 'none';
      toggle.textContent = isHidden ? 'Hide Activity' : 'Show Activity';
      if (isHidden) return;
      closeAllDetailRows(null);
    });

    block.addEventListener('click', e => e.stopPropagation());
  }

  const dateLinks = container.querySelectorAll('.receiving-date-link');
  if (dateLinks && dateLinks.length > 0) {
    dateLinks.forEach(link => {
      link.addEventListener('click', (e) => {
        e.preventDefault();
        e.stopPropagation();

        const dateKey = String(link.getAttribute('data-date') || '').trim();
        if (!dateKey) return;

        const rows = Array.isArray(window.currentReceivingRows) ? window.currentReceivingRows : [];
        const dayRow = rows.find(r => String(r.isoDate || '') === dateKey);
        if (!dayRow) return;

        const detailRow = Array.from(container.querySelectorAll('.receiving-detail-row'))
          .find(row => row.getAttribute('data-date') === dateKey);
        if (!detailRow) return;

        const currentlyVisible = detailRow.style.display !== 'none';
        closeAllDetailRows(null);
        if (currentlyVisible) return;

        detailRow.style.display = '';
        clearReceivingLinkStates();
        setReceivingLinkState(link, true);

        const detailCell = detailRow.querySelector('td');
        if (!detailCell) return;

        const breakdownRows = Array.isArray(dayRow.breakdown) ? dayRow.breakdown : [];
        const detailHtml = breakdownRows.length > 0
          ? `
            <div style="font-size:11px;color:#666;margin-bottom:4px">${breakdownRows.length} piles received material on ${esc(dayRow.dateLabel)}</div>
            <table style="width:100%;font-size:12px;border-collapse:collapse">
              <thead>
                <tr style="background:#f2f2f2">
                  <th style="text-align:left;padding:2px 6px">Pile</th>
                  <th style="text-align:left;padding:2px 6px">Material</th>
                  <th style="text-align:right;padding:2px 6px">Weight (lbs)</th>
                  <th style="text-align:right;padding:2px 6px">Tons</th>
                </tr>
              </thead>
              <tbody>
                ${breakdownRows.map(item => `
                  <tr>
                    <td style="padding:2px 6px">${esc(item.pile || '—')}</td>
                    <td style="padding:2px 6px">${esc(item.material || '')}</td>
                    <td style="padding:2px 6px;text-align:right">${fmtInt(item.weight)}</td>
                    <td style="padding:2px 6px;text-align:right">${fmtTons2(item.tons, 2)}</td>
                  </tr>
                `).join('')}
              </tbody>
            </table>
          `
          : `<div style="font-size:12px;color:#666">No Receiving2 detail for ${esc(dayRow.dateLabel)}.</div>`;

        detailCell.innerHTML = detailHtml;
      });
    });
  }

  const trucksLinks = container.querySelectorAll('.receiving-trucks-link');
  if (trucksLinks && trucksLinks.length > 0) {
    trucksLinks.forEach(link => {
      link.addEventListener('click', (e) => {
        e.preventDefault();
        e.stopPropagation();

        const dateKey = String(link.getAttribute('data-date') || '').trim();
        if (!dateKey) return;

        const rows = Array.isArray(window.currentReceivingRows) ? window.currentReceivingRows : [];
        const dayRow = rows.find(r => String(r.isoDate || '') === dateKey);
        if (!dayRow) return;

        const detailRow = Array.from(container.querySelectorAll('.receiving-truck-detail-row'))
          .find(row => row.getAttribute('data-date') === dateKey);
        if (!detailRow) return;

        const currentlyVisible = detailRow.style.display !== 'none';
        closeAllDetailRows(null);
        if (currentlyVisible) return;

        detailRow.style.display = '';
        clearReceivingLinkStates();
        setReceivingLinkState(link, true);

        const detailCell = detailRow.querySelector('td');
        if (!detailCell) return;

        const truckRows = Array.isArray(dayRow.truckDetails) ? dayRow.truckDetails : [];
        const detailHtml = truckRows.length > 0
          ? `
            <div style="font-size:11px;color:#666;margin-bottom:4px">${fmtInt(dayRow.trucks)} trucks received on ${esc(dayRow.dateLabel)}</div>
            <table style="width:100%;font-size:12px;border-collapse:collapse">
              <thead>
                <tr style="background:#f2f2f2">
                  <th style="text-align:left;padding:2px 6px">Truck #</th>
                  <th style="text-align:left;padding:2px 6px">Ticket ID</th>
                  <th style="text-align:left;padding:2px 6px">Pile</th>
                  <th style="text-align:left;padding:2px 6px">Material</th>
                  <th style="padding:2px 6px;width:52px"></th>
                </tr>
              </thead>
              <tbody>
                ${truckRows.map((item, idx) => {
                  const hasRemarks = !!(item.remarks && item.remarks.length > 0);
                  const hasWeights = !!(item.gross || item.tare || item.net);
                  const rowStyle = item.isStoppedPile ? 'background:#ffebee;' : '';
                  const cellStyle = 'padding:2px 6px;' + (item.isStoppedPile ? 'color:#b71c1c;font-weight:700;' : '');
                  const infoId = 'tki-' + idx;
                  let rows = '<tr style="' + rowStyle + '">'
                    + '<td style="' + cellStyle + '">' + esc(item.truckNumber || '—') + '</td>'
                    + '<td style="' + cellStyle + '">' + esc(item.ticketId || '—') + '</td>'
                    + '<td style="' + cellStyle + '">' + esc(item.pile || '—') + '</td>'
                    + '<td style="' + cellStyle + '">' + esc(item.material || '') + '</td>'
                    + '<td style="padding:2px 4px;white-space:nowrap;text-align:center">'
                    + (hasRemarks ? '<button type="button" class="truck-info-btn" data-target="' + infoId + '-r" title="View remarks" style="border:none;background:none;cursor:pointer;font-size:14px;padding:1px 2px;line-height:1">💬</button>' : '<span style="display:inline-block;font-size:14px;padding:1px 2px;line-height:1;visibility:hidden">💬</span>')
                    + (hasWeights ? '<button type="button" class="truck-info-btn" data-target="' + infoId + '-w" title="View weights" style="border:none;background:none;cursor:pointer;font-size:14px;padding:1px 2px;line-height:1">⚖️</button>' : '<span style="display:inline-block;font-size:14px;padding:1px 2px;line-height:1;visibility:hidden">⚖️</span>')
                    + '</td>'
                    + '</tr>';
                  if (hasRemarks) {
                    rows += '<tr id="' + infoId + '-r" class="truck-info-detail" style="display:none">'
                      + '<td colspan="5" style="padding:2px 8px 5px 24px">'
                      + '<div style="background:#fffde7;border-left:3px solid #f9a825;padding:4px 8px;border-radius:2px;font-size:11px;color:#555;white-space:pre-wrap">' + esc(item.remarks) + '</div>'
                      + '</td></tr>';
                  }
                  if (hasWeights) {
                    rows += '<tr id="' + infoId + '-w" class="truck-info-detail" style="display:none">'
                      + '<td colspan="5" style="padding:2px 8px 5px 24px">'
                      + '<div style="background:#e8f5e9;border-left:3px solid #43a047;padding:4px 8px;border-radius:2px;font-size:11px;color:#333">'
                      + '<strong>Gross:</strong> ' + esc(item.gross || '—') + '&nbsp;&nbsp;<strong>Tare:</strong> ' + esc(item.tare || '—') + '&nbsp;&nbsp;<strong>Net:</strong> ' + esc(item.net || '—')
                      + '</div></td></tr>';
                  }
                  return rows;
                }).join('')}
              </tbody>
            </table>
          `
          : `<div style="font-size:12px;color:#666">No Receiving3 detail for ${esc(dayRow.dateLabel)}.</div>`;

        detailCell.innerHTML = detailHtml;

        detailCell.querySelectorAll('.truck-info-btn').forEach(btn => {
          btn.addEventListener('click', e => {
            e.stopPropagation();
            const targetId = btn.getAttribute('data-target');
            const targetRow = detailCell.querySelector('#' + targetId);
            if (!targetRow) return;
            const isOpen = targetRow.style.display !== 'none';
            detailCell.querySelectorAll('.truck-info-detail').forEach(r => { r.style.display = 'none'; });
            if (!isOpen) targetRow.style.display = '';
          });
        });
      });
    });
  }
}

function wireRailcarPopupEvents(container, marker) {
  const railcarMonthBtn = container.querySelector('#railcarMonthBtn');
  if (railcarMonthBtn) {
    railcarMonthBtn.addEventListener('click', async (e) => {
      e.stopPropagation(); e.preventDefault();
      railcarMonthBtn.textContent = 'Loading…';
      railcarMonthBtn.disabled = true;
      if (!railcarHistoryMonths) railcarHistoryMonths = await discoverHistoryMonths('Railcars');
      const payload = await fetchRailcarSummary(false, railcarCurrentPeriod);
      marker.setPopupContent(unescapeAngles(renderRailcarPopup(payload)));
      setTimeout(() => { const el = marker.getPopup()?.getElement(); if (el) wireRailcarPopupEvents(el, marker); }, 0);
    });
  }
  const railcarCalToggle = container.querySelector('#railcarMonthCalendarToggle');
  if (railcarCalToggle) {
    railcarCalToggle.addEventListener('click', (e) => {
      e.preventDefault(); e.stopPropagation();
      const cal = container.querySelector('#railcarMonthCalendar');
      if (cal) cal.style.display = cal.style.display === 'none' ? '' : 'none';
    });
  }
  container.querySelectorAll('.railcarMonthCell').forEach(cell => {
    cell.addEventListener('click', async (e) => {
      e.preventDefault(); e.stopPropagation();
      const val = cell.getAttribute('data-value');
      railcarCurrentPeriod = val === 'current' ? null : { year: +val.split('-')[0], month: +val.split('-')[1] };
      const payload = await fetchRailcarSummary(false, railcarCurrentPeriod);
      marker.setPopupContent(unescapeAngles(renderRailcarPopup(payload)));
      setTimeout(() => { const el = marker.getPopup()?.getElement(); if (el) wireRailcarPopupEvents(el, marker); }, 0);
    });
  });
  const railcarResetBtn = container.querySelector('#railcarMonthResetBtn');
  if (railcarResetBtn) {
    railcarResetBtn.addEventListener('click', async (e) => {
      e.preventDefault(); e.stopPropagation();
      railcarCurrentPeriod = null;
      const payload = await fetchRailcarSummary(false, null);
      marker.setPopupContent(unescapeAngles(renderRailcarPopup(payload)));
      setTimeout(() => { const el = marker.getPopup()?.getElement(); if (el) wireRailcarPopupEvents(el, marker); }, 0);
    });
  }

  const pairs = [
    { btn: container.querySelector('#railcarOnsiteToggle'),  section: container.querySelector('#railcarOnsiteSection') },
    { btn: container.querySelector('#railcarReleasedToggle'), section: container.querySelector('#railcarReleasedSection') }
  ];

  function closePanel(p) {
    if (!p.btn || !p.section) return;
    p.section.style.display = 'none';
    p.btn.textContent = p.btn.textContent.replace(/^Hide\s/, '');
  }

  pairs.forEach(active => {
    if (!active.btn || !active.section) return;
    active.btn.addEventListener('click', () => {
      const isOpen = active.section.style.display !== 'none';
      if (!isOpen) pairs.forEach(other => { if (other !== active) closePanel(other); });
      active.section.style.display = isOpen ? 'none' : '';
      active.btn.textContent = isOpen
        ? active.btn.textContent.replace(/^Hide\s/, '')
        : 'Hide ' + active.btn.textContent;
    });
  });

  container.querySelectorAll('.railcar-photo-btn').forEach(btn => {
    btn.addEventListener('click', e => {
      e.stopPropagation();
      const car = btn.getAttribute('data-car');
      const month = btn.getAttribute('data-month');
      if (car) showRailcarPhoto(car, month);
    });
  });

  // Instant offset tooltip for [data-offset] cells
  const tip = document.createElement('div');
  tip.style.cssText = 'position:fixed;z-index:99999;background:rgba(30,30,30,0.92);color:#fff;font-size:11px;padding:5px 9px;border-radius:3px;pointer-events:none;white-space:nowrap;display:none;box-shadow:0 2px 6px rgba(0,0,0,0.3);';
  document.body.appendChild(tip);

  marker.once('popupclose', () => { if (tip.parentNode) tip.parentNode.removeChild(tip); });

  container.querySelectorAll('[data-offset]').forEach(cell => {
    cell.addEventListener('mouseenter', e => {
      tip.textContent = cell.getAttribute('data-offset');
      tip.style.display = 'block';
      tip.style.left = (e.clientX + 14) + 'px';
      tip.style.top  = (e.clientY - 10) + 'px';
    });
    cell.addEventListener('mousemove', e => {
      tip.style.left = (e.clientX + 14) + 'px';
      tip.style.top  = (e.clientY - 10) + 'px';
    });
    cell.addEventListener('mouseleave', () => { tip.style.display = 'none'; });
  });
}

function wireBreakingPopupEvents(container, marker) {
  const breakingMonthBtn = container.querySelector('#breakingMonthBtn');
  if (breakingMonthBtn) {
    breakingMonthBtn.addEventListener('click', async (e) => {
      e.stopPropagation(); e.preventDefault();
      breakingMonthBtn.textContent = 'Loading…';
      breakingMonthBtn.disabled = true;
      if (!breakingHistoryMonths) breakingHistoryMonths = await discoverHistoryMonthsForYearlySheet('BreakingHistory');
      const payload = await fetchBreakingTotals(false, breakingCurrentPeriod);
      marker.setPopupContent(unescapeAngles(renderBreakingPopup(payload, allMarkersData, stockIndexGlobal)));
      setTimeout(() => { const el = marker.getPopup()?.getElement(); if (el) wireBreakingPopupEvents(el, marker); }, 0);
    });
  }
  const breakingCalToggle = container.querySelector('#breakingMonthCalendarToggle');
  if (breakingCalToggle) {
    breakingCalToggle.addEventListener('click', (e) => {
      e.preventDefault(); e.stopPropagation();
      const cal = container.querySelector('#breakingMonthCalendar');
      if (cal) cal.style.display = cal.style.display === 'none' ? '' : 'none';
    });
  }
  container.querySelectorAll('.breakingMonthCell').forEach(cell => {
    cell.addEventListener('click', async (e) => {
      e.preventDefault(); e.stopPropagation();
      const val = cell.getAttribute('data-value');
      breakingCurrentPeriod = val === 'current' ? null : { year: +val.split('-')[0], month: +val.split('-')[1] };
      const payload = await fetchBreakingTotals(false, breakingCurrentPeriod);
      marker.setPopupContent(unescapeAngles(renderBreakingPopup(payload, allMarkersData, stockIndexGlobal)));
      setTimeout(() => { const el = marker.getPopup()?.getElement(); if (el) wireBreakingPopupEvents(el, marker); }, 0);
    });
  });
  const breakingResetBtn = container.querySelector('#breakingMonthResetBtn');
  if (breakingResetBtn) {
    breakingResetBtn.addEventListener('click', async (e) => {
      e.preventDefault(); e.stopPropagation();
      breakingCurrentPeriod = null;
      const payload = await fetchBreakingTotals(false, null);
      marker.setPopupContent(unescapeAngles(renderBreakingPopup(payload, allMarkersData, stockIndexGlobal)));
      setTimeout(() => { const el = marker.getPopup()?.getElement(); if (el) wireBreakingPopupEvents(el, marker); }, 0);
    });
  }

  // Activity
  const activityToggle = container.querySelector('#breakingToggle');
  const activityBlock  = container.querySelector('#breakingActivity');

  // NEW: Download button
  const dlBtn = container.querySelector('#breakingDownload');

  // Unprep top/bottom buttons + section
  const unprepTopBtn = container.querySelector('#unprepToggleTop');
  const unprepBtmBtn = container.querySelector('#unprepToggleBottom');
  const unprepDiv    = container.querySelector('#unprepSection');

  // ---- helpers ----
  function isActivityShown() { return activityBlock && activityBlock.style.display !== 'none'; }
  function isUnprepShown()   { return unprepDiv && unprepDiv.style.display === 'block'; }

  function setDownloadVisibility() {
    if (!dlBtn) return;
    // Mirror Burning: show download only when Activity is visible and Unprep is not open
    dlBtn.style.display = (isActivityShown() && !isUnprepShown()) ? '' : 'none';
  }

  function setUnprepState(show) {
    if (!unprepDiv) return;
    unprepDiv.style.display = show ? 'block' : 'none';
    if (unprepTopBtn) {
      unprepTopBtn.textContent = show ? 'Hide Unprep' : 'Show Unprep';
      unprepTopBtn.style.display = show ? 'none' : ''; // hide top button when open
    }
    if (unprepBtmBtn) {
      unprepBtmBtn.textContent = show ? 'Hide Unprep' : 'Show Unprep';
    }
    // Hide Activity button while Unprep is open (mirroring Burning)
    if (activityToggle) activityToggle.style.display = show ? 'none' : '';
    setDownloadVisibility(); // keep download in sync
  }

  // Initialize Unprep
  if (unprepDiv && (unprepTopBtn || unprepBtmBtn)) {
    setUnprepState(unprepDiv.style.display === 'block');
    const onUnprepClick = () => {
      const showing = unprepDiv.style.display === 'block';
      setUnprepState(!showing);
    };
    if (unprepTopBtn) unprepTopBtn.addEventListener('click', onUnprepClick);
    if (unprepBtmBtn) unprepBtmBtn.addEventListener('click', onUnprepClick);
  }

  // Activity toggle (mirrors Burning)
  if (activityToggle && activityBlock) {
    if (activityBlock.style.display === '') activityBlock.style.display = 'none';
    activityToggle.addEventListener('click', e => {
      e.preventDefault(); e.stopPropagation();
      const hidden = activityBlock.style.display === 'none';
      if (hidden) {
        activityBlock.style.display = '';
        activityToggle.textContent = 'Hide Activity';
        if (unprepTopBtn) unprepTopBtn.style.display = 'none';
      } else {
        activityBlock.style.display = 'none';
        activityToggle.textContent = 'Show Activity';
        if (unprepTopBtn && !isUnprepShown()) unprepTopBtn.style.display = '';
      }
      setDownloadVisibility();
    });
    activityBlock.addEventListener('click', e => e.stopPropagation());
  } else {
    // If there's no activity section, make sure download is hidden
    if (dlBtn) dlBtn.style.display = 'none';
  }

  // NEW: Download handler for Breaking (current month only)
  if (dlBtn) {
    dlBtn.addEventListener('click', e => {
      e.preventDefault(); e.stopPropagation();
      const rows = window.currentBreakingSheetRows || [];
      const totals = window.currentBreakingTotals || {};
      const exportRows = rows.map(r => ({
        Date: r.dateLabel || '',
        Commodity: r.material || '',
        Net: r.netLbs ?? '',
        'Net Tons': r.netTons ?? '',
        'From Pile #': r.from || '',
        'To Pile #': r.to || ''
      }));
      // Append totals row
      exportRows.push({
        Date: 'TOTAL',
        Commodity: '',
        Net: totals.totalLbs ?? '',
        'Net Tons': totals.totalTons ?? '',
        'From Pile #': '',
        'To Pile #': ''
      });
      downloadCsvFromRows('Breaking.csv', exportRows, ['Date', 'Commodity', 'Net', 'Net Tons', 'From Pile #', 'To Pile #']);
    });
  }

  // Initial visibility
  setDownloadVisibility();
}

function wireBurningPopupEvents(container, marker) {
  const burningMonthBtn = container.querySelector('#burningMonthBtn');
  if (burningMonthBtn) {
    burningMonthBtn.addEventListener('click', async (e) => {
      e.stopPropagation(); e.preventDefault();
      burningMonthBtn.textContent = 'Loading…';
      burningMonthBtn.disabled = true;
      if (!burningHistoryMonths) burningHistoryMonths = await discoverHistoryMonthsForYearlySheet('BurningHistory');
      const payload = await fetchBurningTotals(false, burningCurrentPeriod);
      marker.setPopupContent(unescapeAngles(renderBurningPopup(payload, allMarkersData, stockIndexGlobal)));
      setTimeout(() => { const el = marker.getPopup()?.getElement(); if (el) wireBurningPopupEvents(el, marker); }, 0);
    });
  }
  const burningCalToggle = container.querySelector('#burningMonthCalendarToggle');
  if (burningCalToggle) {
    burningCalToggle.addEventListener('click', (e) => {
      e.preventDefault(); e.stopPropagation();
      const cal = container.querySelector('#burningMonthCalendar');
      if (cal) cal.style.display = cal.style.display === 'none' ? '' : 'none';
    });
  }
  container.querySelectorAll('.burningMonthCell').forEach(cell => {
    cell.addEventListener('click', async (e) => {
      e.preventDefault(); e.stopPropagation();
      const val = cell.getAttribute('data-value');
      burningCurrentPeriod = val === 'current' ? null : { year: +val.split('-')[0], month: +val.split('-')[1] };
      const payload = await fetchBurningTotals(false, burningCurrentPeriod);
      marker.setPopupContent(unescapeAngles(renderBurningPopup(payload, allMarkersData, stockIndexGlobal)));
      setTimeout(() => { const el = marker.getPopup()?.getElement(); if (el) wireBurningPopupEvents(el, marker); }, 0);
    });
  });
  const burningResetBtn = container.querySelector('#burningMonthResetBtn');
  if (burningResetBtn) {
    burningResetBtn.addEventListener('click', async (e) => {
      e.preventDefault(); e.stopPropagation();
      burningCurrentPeriod = null;
      const payload = await fetchBurningTotals(false, null);
      marker.setPopupContent(unescapeAngles(renderBurningPopup(payload, allMarkersData, stockIndexGlobal)));
      setTimeout(() => { const el = marker.getPopup()?.getElement(); if (el) wireBurningPopupEvents(el, marker); }, 0);
    });
  }

  const refresh = container.querySelector('#burningRefresh');

  // Activity pieces
  const toggle = container.querySelector('#burningToggle');        // Activity toggle button
  const block  = container.querySelector('#burningActivity');      // Activity section

  // Coils pieces (TOP button in control row, BOTTOM button in coils footer)
  const coilsTopBtn = container.querySelector('#coilsToggleTop');
  const coilsBtmBtn = container.querySelector('#coilsToggleBottom');
  const coilsDiv    = container.querySelector('#coilsSection');

  // Download button
  const dlBtn = container.querySelector('#burningDownload');

  // ---------- Helpers ----------
  function isActivityShown() {
    return block && block.style.display !== 'none';
  }
  function isCoilsShown() {
    return coilsDiv && coilsDiv.style.display === 'block';
  }
  function setDownloadVisibility() {
    // Show download only when Activity is visible and Coils is not open
    if (!dlBtn) return;
    dlBtn.style.display = (isActivityShown() && !isCoilsShown()) ? '' : 'none';
  }

  // ---------- Coils toggle logic (sync top & bottom buttons) ----------
  function setCoilsState(show) {
    if (!coilsDiv) return;

    // Section visibility
    coilsDiv.style.display = show ? 'block' : 'none';

    // Button labels & visibility
    if (coilsTopBtn) {
      coilsTopBtn.textContent = show ? 'Hide Coils' : 'Show Coils';
      // Hide the TOP button while coils section is open (the bottom one is available)
      coilsTopBtn.style.display = show ? 'none' : '';
    }
    if (coilsBtmBtn) {
      coilsBtmBtn.textContent = show ? 'Hide Coils' : 'Show Coils';
      // Bottom button lives inside the section; appears automatically when open
    }

    // Hide the Activity toggle while Coils are open (matches your prior UX)
    if (toggle) toggle.style.display = show ? 'none' : '';

    // Update Download button visibility based on new state
    setDownloadVisibility();
  }

  if (coilsDiv && (coilsTopBtn || coilsBtmBtn)) {
    // Initialize coils state (template starts hidden)
    setCoilsState(coilsDiv.style.display === 'block');

    const onCoilsClick = () => {
      const showing = coilsDiv.style.display === 'block';
      setCoilsState(!showing);
    };
    if (coilsTopBtn) coilsTopBtn.addEventListener('click', onCoilsClick);
    if (coilsBtmBtn) coilsBtmBtn.addEventListener('click', onCoilsClick);
  }

  // ---------- Activity toggle logic ----------
  if (toggle && block) {
    // Ensure initial state is hidden (as per your template), then set initial download visibility
    if (block.style.display === '') block.style.display = 'none';
    setDownloadVisibility();

    toggle.addEventListener('click', e => {
      e.preventDefault(); e.stopPropagation();
      const hidden = block.style.display === 'none';

      if (hidden) {
        // Show Activity, hide coils TOP button to avoid competing toggles
        block.style.display = '';
        toggle.textContent = 'Hide Activity';
        if (coilsTopBtn) coilsTopBtn.style.display = 'none';
      } else {
        // Hide Activity, re-show coils TOP button only if coils are not open
        block.style.display = 'none';
        toggle.textContent = 'Show Activity';
        if (coilsTopBtn && !isCoilsShown()) coilsTopBtn.style.display = '';
      }

      // Update Download visibility after toggling Activity
      setDownloadVisibility();
    });

    block.addEventListener('click', e => e.stopPropagation());
  } else {
    // If there's no activity block/toggle, make sure download is hidden
    if (dlBtn) dlBtn.style.display = 'none';
  }

  // ---------- Refresh logic ----------
  if (refresh) {
    refresh.addEventListener('click', async e => {
      e.preventDefault(); e.stopPropagation();
      try {
        const data = await fetchBurningTotals(true);
        const encoded = renderBurningPopup(data, allMarkersData, stockIndexGlobal);
        const decoded = unescapeAngles(encoded);
        marker.setPopupContent(decoded);

        setTimeout(() => {
          const el = marker.getPopup()?.getElement();
          if (el) wireBurningPopupEvents(el, marker);
        }, 0);
      } catch (err) { console.error(err); }
    });
  }

  // ---------- Download logic ----------
  if (dlBtn) {
    dlBtn.addEventListener('click', e => {
      e.preventDefault(); e.stopPropagation();
      const rows = window.currentBurningSheetRows || [];
      const totals = window.currentBurningTotals || {};
      const exportRows = rows.map(r => ({
        Date: r.dateLabel || '',
        'Net(lbs)': r.netLbs ?? '',
        From: r.from || '',
        To: r.to || '',
        NetTons: r.netTons ?? '',
        '# of Cuts': r.cuts ?? '',
        'Billable Tons': r.billableTons ?? ''
      }));
      // Append totals row
      exportRows.push({
        Date: 'TOTAL',
        'Net(lbs)': totals.netLbs ?? '',
        From: '',
        To: '',
        NetTons: totals.netTons ?? '',
        '# of Cuts': totals.cuts ?? '',
        'Billable Tons': totals.billableTons ?? ''
      });
      downloadCsvFromRows('Burning.csv', exportRows, ['Date', 'Net(lbs)', 'From', 'To', 'NetTons', '# of Cuts', 'Billable Tons']);
    });
  }
}

// 6) clickable marker --------------------------------------
const burningArea = L.circleMarker(burningLatLng, {
  radius: 18,
  color: 'rgba(255,255,0,0.01)',
  fillColor: 'rgba(0,0,0,0.01)',
  fillOpacity: 0.01,
  weight: 12
}).addTo(map);
window.burningArea = burningArea;

burningArea.bindPopup('', { maxWidth: 420, autopan: false });
burningArea.bindTooltip('Burning Station', { permanent: false, direction: 'top' });

burningArea.on('popupopen', async () => {
  burningArea.setPopupContent(`&lt;div style="${POPUP_CONTAINER_STYLE}"&gt;Loading…&lt;/div&gt;`);
  try {
    const payload = await fetchBurningTotals(false, burningCurrentPeriod);
    const encoded = renderBurningPopup(payload, allMarkersData, stockIndexGlobal);
    const decoded = unescapeAngles(encoded);
    burningArea.setPopupContent(decoded);

    setTimeout(() => {
      const el = burningArea.getPopup()?.getElement();
      if (el) wireBurningPopupEvents(el, burningArea);
    }, 0);

  } catch (err) {
    console.error(err);
    burningArea.setPopupContent('&lt;b&gt;Burning Station&lt;/b&gt;&lt;div style="color:#c00"&gt;Failed to load Production.xlsx Burning sheet.&lt;/div&gt;');
  }
});

const breakingArea = L.circleMarker(breakingLatLng, {
  radius: 18,
  color: 'rgba(255,255,0,0.01)',
  fillColor: 'rgba(0,0,0,0.01)',
  fillOpacity: 0.01,
  weight: 12
}).addTo(map);
window.breakingArea = breakingArea;

breakingArea.bindPopup('', { maxWidth: 420, autopan: false });
breakingArea.bindTooltip('Breaking Pit', { permanent: false, direction: 'top' });

breakingArea.on('popupopen', async () => {
  breakingArea.setPopupContent(`<div style="${POPUP_CONTAINER_STYLE}">Loading…</div>`);
  try {
    const payload = await fetchBreakingTotals(false, breakingCurrentPeriod);
    const encoded = renderBreakingPopup(payload, allMarkersData, stockIndexGlobal);
    const decoded = unescapeAngles(encoded);
    breakingArea.setPopupContent(decoded);

    setTimeout(() => {
      const el = breakingArea.getPopup()?.getElement();
      if (el) wireBreakingPopupEvents(el, breakingArea);
    }, 0);
  } catch (err) {
    console.error(err);
    breakingArea.setPopupContent('<b>Breaking Pit</b><div style="color:#c00">Failed to load Production.xlsx Breaking sheet.</div>');
  }
});

const bucketLoadingArea = L.circleMarker(bucketLoadingLatLng, {
  radius: 18,
  color: 'rgba(255,255,0,0.01)',
  fillColor: 'rgba(0,0,0,0.01)',
  fillOpacity: 0.01,
  weight: 12
}).addTo(map);
window.bucketLoadingArea = bucketLoadingArea;

bucketLoadingArea.bindPopup('', { maxWidth: 420, autopan: false });
bucketLoadingArea.bindTooltip('Bucket Loading', { permanent: false, direction: 'top' });

bucketLoadingArea.on('popupopen', async () => {
  bucketLoadingArea.setPopupContent(`<div style="${POPUP_CONTAINER_STYLE}">Loading…</div>`);
  try {
    const payload = await fetchBucketLoadingConsumption(false, bucketCurrentPeriod);
    const encoded = renderBucketLoadingPopup(payload);
    const decoded = unescapeAngles(encoded);
    bucketLoadingArea.setPopupContent(decoded);

    setTimeout(() => {
      const el = bucketLoadingArea.getPopup()?.getElement();
      if (el) wireBucketLoadingPopupEvents(el, bucketLoadingArea);
    }, 0);
  } catch (err) {
    console.error(err);
    bucketLoadingArea.setPopupContent('<b>Bucket Loading</b><div style="color:#c00">Failed to load Consumption1 sheet.</div>');
  }
});

const receivingArea = L.circleMarker(receivingLatLng, {
  radius: 18,
  color: 'rgba(255,255,0,0.01)',
  fillColor: 'rgba(0,0,0,0.01)',
  fillOpacity: 0.01,
  weight: 12
}).addTo(map);
window.receivingArea = receivingArea;

receivingArea.bindPopup('', { maxWidth: 420, autopan: false });
receivingArea.bindTooltip('Truck Scales', { permanent: false, direction: 'top' });

receivingArea.on('popupopen', async () => {
  receivingArea.setPopupContent(`<div style="${POPUP_CONTAINER_STYLE}">Loading…</div>`);
  try {
    const payload = await fetchReceivingSummary(false, receivingCurrentPeriod);
    const encoded = renderReceivingPopup(payload);
    const decoded = unescapeAngles(encoded);
    receivingArea.setPopupContent(decoded);

    setTimeout(() => {
      const el = receivingArea.getPopup()?.getElement();
      if (el) wireReceivingPopupEvents(el, receivingArea);
    }, 0);
  } catch (err) {
    console.error(err);
    receivingArea.setPopupContent('<b>Truck Scales</b><div style="color:#c00">Failed to load Receiving1 sheet.</div>');
  }
});

/* ===================================================================
 RAILROAD PANEL
=================================================================== */
const railcarArea = L.circleMarker(railcarLatLng, {
  radius: 18,
  color: 'rgba(255,255,0,0.01)',
  fillColor: 'rgba(0,0,0,0.01)',
  fillOpacity: 0.01,
  weight: 12
}).addTo(map);
window.railcarArea = railcarArea;

railcarArea.bindPopup('', { maxWidth: 720, autopan: false });
railcarArea.bindTooltip('Railroad', { permanent: false, direction: 'top' });

railcarArea.on('popupopen', async () => {
  railcarArea.setPopupContent(`<div style="min-width:500px">Loading…</div>`);
  try {
    if (!railcarHistoryMonths) railcarHistoryMonths = await discoverHistoryMonths('Railcars');
    const payload = await fetchRailcarSummary(false, railcarCurrentPeriod);
    const encoded = renderRailcarPopup(payload);
    const decoded = unescapeAngles(encoded);
    railcarArea.setPopupContent(decoded);

    setTimeout(() => {
      const el = railcarArea.getPopup()?.getElement();
      if (el) wireRailcarPopupEvents(el, railcarArea);
    }, 0);
  } catch (err) {
    console.error(err);
    railcarArea.setPopupContent('<b>Railroad</b><div style="color:#c00">Failed to load Railcars sheet.</div>');
  }
});

/* ===================================================================
 LOAD MARKERS + ENRICH POPUPS
=================================================================== */
Promise.all([
  fetch('markers.json').then(r => r.json()),
  fetchLatestInventoryCsv()
]).then(([markers, stockPayload]) => {
  allMarkersData = markers;                 // save globally
  rebuildStoppedPileCodes(allMarkersData);  // keep stopped pile set in sync with markers.json
  stockIndexGlobal = stockPayload.stock || {};
  
  const unknownTypes = new Set();

  // Helpers
  function formatAgeYM(totalMonths) {
    if (typeof totalMonths !== 'number' || !isFinite(totalMonths) || totalMonths < 0) return '';
    const years = Math.floor(totalMonths / 12);
    const months = totalMonths % 12;
    const yPart = years > 0 ? `${years} ${years === 1 ? 'year' : 'years'}` : '';
    const mPart = months > 0 ? `${months} ${months === 1 ? 'month' : 'months'}` : (years === 0 ? '0 months' : '');
    return yPart && mPart ? `${yPart} ${mPart}` : (yPart || mPart);
  }

function renderPopupHtml(marker, s) {
  if (!s) return `<b>${marker.name}</b>`;

  const inv = (typeof s.operating_inventory_lbs === 'number')
    ? s.operating_inventory_lbs
    : Number(s.operating_inventory_lbs || 0);

  const invText = Number.isFinite(inv)
    ? inv.toLocaleString('en-US', { maximumFractionDigits: 0 }) + ' lbs'
    : '—';

  const invStyle = invAlertStyle(inv);
  const matStyle = invStyle; // material follows the same rule
  const lzStyle  = lastZeroColor(s.last_zero_date);

  return `
    <div style="min-width:220px">
      <div style="font-weight:700;margin-bottom:4px">${marker.name}</div>
      <table style="font-size:12px;line-height:1.3">
        <tr>
          <td style="padding-right:8px;color:#666">Material:</td>
          <td style="${matStyle}">${s.material ?? ''}</td>
        </tr>
        <tr>
          <td style="padding-right:8px;color:#666">Inventory:</td>
          <td style="${invStyle}">${invText}</td>
        </tr>
        <tr>
          <td style="padding-right:8px;color:#666">Last Zero Date:</td>
          <td style="${lzStyle}">${s.last_zero_date ?? ''}</td>
        </tr>
      </table>
    </div>
  `;
}

  // Create markers
  markers.forEach(marker => {
    const cfg = markerConfig[marker.type];
    if (!cfg) { unknownTypes.add(marker.type); return; }
    const m = L.marker([marker.lat, marker.lng], { icon: cfg.icon });
    const code = extractPileCode(marker.name);
    const s = code ? stockIndexGlobal[code] : null;
    m.bindPopup('', { maxWidth: 320, autopan: false });
    m.on('popupopen', () => m.setPopupContent(renderPopupHtml(marker, s)));
    cfg.layer.addLayer(m);
    m._group = cfg.layer;
    marker._leaflet = m;
    if (marker.cell) {
      (Array.isArray(marker.cell) ? marker.cell : [marker.cell]).forEach(c => {
        if (loadCells[c]) loadCells[c].addLayer(m);
      });
    }
  });

  // Past Due list (> 6 months)
  const pastDue = [];
  const seenCodes = new Set();
  const exemptTypes = new Set(["Coils", "Breaking", "Unbreakable", "Alloys"]);
function isPastDueExempt(marker, stockIndex) {
  if (!marker) return true;
  if (!exemptTypes.has(marker.type)) return false;
  if (marker.type === "Alloys") {
    const name = String(marker.name || "").toLowerCase();
    const code = extractPileCode(marker.name);
    const s = code ? stockIndex[code] : null;
    const material = String((s && s.material) || "").toLowerCase();
    const isMoOx =
      name.includes("molybdenum oxide") ||
      material.includes("molybdenum oxide");
    return !isMoOx;
  }
  return true;
}

  markers.forEach(marker => {
    if (isPastDueExempt(marker, stockIndexGlobal)) return;
    const code = extractPileCode(marker.name);
    if (!code) return;
    const s = stockIndexGlobal[code];
    if (!s) return;
    const dt = parseMDY(s.last_zero_date);
    if (!dt) return;
    const mAge = monthsDiff(dt, new Date());
    if (mAge > 6 && !seenCodes.has(code)) {
      pastDue.push({
        code: code,
        name: marker.name,
        rawType: marker.type,
        material: s.material || '',
        lastZero: s.last_zero_date,
        ageMonths: mAge,
        ageLabel: formatAgeYM(mAge),
        invLbs: s.operating_inventory_lbs,
        marker: marker._leaflet
      });
      seenCodes.add(code);
    }
  });
  pastDue.sort((a, b) => (b.ageMonths - a.ageMonths) || a.code.localeCompare(b.code));
  window.pastDue = pastDue;

  /* ===================================================================
 SEARCH PANEL HELPER FUNCTIONS
=================================================================== */

  function getVisiblePiles() {
    if (searchMode === 'pastDue') {
      return pastDue.map(p => ({
        code: p.code,
        name: p.name,
        material: p.material,
        marker: p.marker,
        invLbs: p.invLbs,
        lastZero: p.lastZero,
        ageLabel: p.ageLabel,
        type: p.rawType
      }));
    }

    return markers.map(m => {
      const code = extractPileCode(m.name);
      const s = code ? stockIndexGlobal[code] : null;
      const lz = s?.last_zero_date ?? '';
      const dt = parseMDY(lz);
      const ageMonths = dt ? monthsDiff(dt, new Date()) : null;
      return {
        code,
        name: m.name,
        material: s?.material ?? '',
        marker: m._leaflet,
        type: m.type,
        invLbs: (s && typeof s.operating_inventory_lbs === 'number') ? s.operating_inventory_lbs : null,
        lastZero: lz,
        ageLabel: (ageMonths !== null && ageMonths >= 0) ? formatAgeYM(ageMonths) : ''
      };
    });
  }

  // Helper function to ping piles and their related stations
  function pingPileWithStation(pile) {
    if (pile?.marker) {
      pingMarker(pile.marker);
    }
    
    // Ping station markers for special pile types
    if (pile?.type === 'Coils' && window.burningArea) {
      pingMarker(window.burningArea);
    } else if ((pile?.type === 'Breaking' || pile?.type === 'Unbreakable') && window.breakingArea) {
      pingMarker(window.breakingArea);
    }
  }

  /* ===================================================================
 SEARCH MODAL - Magnifying Glass in Bottom Right
=================================================================== */

  function createSearchModal() {
    if (isSearchModalOpen) {
      return null;
    }
    isSearchModalOpen = true;

    const searchButton = document.getElementById('searchPanelToggle');
    if (searchButton) {
      searchButton.style.display = 'none';
    }

    // Backdrop for closing modal on click outside
    const backdrop = document.createElement('div');
    backdrop.id = 'searchModalBackdrop';
    backdrop.style.cssText = `
      position: fixed;
      top: 0;
      left: 0;
      width: 100%;
      height: 100%;
      z-index: 899;
      pointer-events: none;
    `;

    // Modal container - positioned in bottom right
    const modal = document.createElement('div');
    modal.id = 'searchModal';
    modal.style.cssText = `
      position: fixed;
      bottom: 8px;
      right: 8px;
      background: rgba(255, 255, 255, 0.98);
      width: 500px;
      max-height: 45vh;
      border-radius: 8px;
      box-shadow: 0 2px 8px rgba(0, 0, 0, 0.2);
      display: flex;
      flex-direction: column;
      font-family: system-ui, -apple-system, Segoe UI, Roboto, Arial, sans-serif;
      font-size: 12px;
      z-index: 900;
      pointer-events: auto;
    `;

    // Header
    const header = document.createElement('div');
    header.style.cssText = `
      display: flex;
      align-items: center;
      justify-content: space-between;
      border-bottom: 1px solid #ddd;
      padding: 12px 16px;
      font-weight: 700;
      font-size: 14px;
    `;
    header.innerHTML = `
      <span>Search Piles <span id="pileCount" style="color: #666; font-weight: normal;">(—)</span></span>
      <button id="searchModalClose" style="
        background: #f0f0f0;
        border: 1px solid #ccc;
        border-radius: 4px;
        font-size: 24px;
        color: #333;
        cursor: pointer;
        padding: 2px 8px;
        width: auto;
        height: auto;
        display: flex;
        align-items: center;
        justify-content: center;
        font-weight: bold;
        transition: all 0.2s ease;
        line-height: 1;
      " onmouseover="this.style.background='#e0e0e0'; this.style.color='#000';" onmouseout="this.style.background='#f0f0f0'; this.style.color='#333';">×</button>
    `;

    // Mode buttons
    const modeContainer = document.createElement('div');
    modeContainer.style.cssText = `
      display: flex;
      gap: 8px;
      padding: 12px 16px;
      border-bottom: 1px solid #ddd;
      flex-wrap: wrap;
    `;
    modeContainer.innerHTML = `
      <button id="modeAll" style="
        padding: 6px 12px;
        border: 1px solid #ccc;
        border-radius: 4px;
        background: #f8f8f8;
        cursor: pointer;
        font-weight: 700;
      ">All</button>
      <button id="modePast" style="
        padding: 6px 12px;
        border: 1px solid #ccc;
        border-radius: 4px;
        background: #f8f8f8;
        cursor: pointer;
      ">Past Due (${pastDue.length})</button>
      <button id="pdExport" style="
        padding: 6px 12px;
        border: 1px solid #ccc;
        border-radius: 4px;
        background: #f8f8f8;
        cursor: pointer;
        margin-left: auto;
        display: none;
      ">Export</button>
    `;

    // Search input
    const inputContainer = document.createElement('div');
    inputContainer.id = 'inputContainer';
    inputContainer.style.cssText = `
      padding: 12px 16px;
      border-bottom: 1px solid #ddd;
    `;
    inputContainer.innerHTML = `
      <input id="pileSearch"
        placeholder="Filter piles..."
        style="
          width: 100%;
          padding: 8px 12px;
          border: 1px solid #ddd;
          border-radius: 4px;
          box-sizing: border-box;
          font-size: 12px;
        " />
    `;

    // Pile list container
    const listContainer = document.createElement('div');
    listContainer.style.cssText = `
      flex: 1;
      overflow-y: auto;
      padding: 8px 0;
    `;

    const pileList = document.createElement('ul');
    pileList.id = 'pileList';
    pileList.style.cssText = `
      list-style: none;
      padding: 0;
      margin: 0;
    `;

    listContainer.appendChild(pileList);

    // Assemble modal
    modal.appendChild(header);
    modal.appendChild(modeContainer);
    modal.appendChild(inputContainer);
    modal.appendChild(listContainer);

    // Append both modal and backdrop to body (modal is positioned fixed independently)
    document.body.appendChild(backdrop);
    document.body.appendChild(modal);

    // State
    let activeIndex = -1;
    let currentFilteredRows = []; // Store filtered rows for keyboard navigation matching
    const input = inputContainer.querySelector('#pileSearch');
    const btnAll = modeContainer.querySelector('#modeAll');
    const btnPast = modeContainer.querySelector('#modePast');
    const btnExport = modeContainer.querySelector('#pdExport');
    const closeBtn = header.querySelector('#searchModalClose');

    // Set modal to be focusable for keyboard events
    modal.tabIndex = 0;
    modal.focus();

    // Render pile list
    function renderList() {
      const term = input.value.trim().toLowerCase();
      pileList.innerHTML = '';

      let rows = getVisiblePiles();
      if (searchMode === 'all' && term) {
        rows = rows.filter(p =>
          (p.code && p.code.toLowerCase().includes(term)) ||
          (p.name && p.name.toLowerCase().includes(term)) ||
          (p.material && p.material.toLowerCase().includes(term))
        );
      }

      // Store filtered rows for keyboard navigation
      currentFilteredRows = rows;

      // Update pile count in header
      const pileCountEl = document.getElementById('pileCount');
      if (pileCountEl) {
        pileCountEl.textContent = `(${rows.length})`;
      }

      rows.forEach((p, idx) => {
        const li = document.createElement('li');
        li.style.cssText = `
          padding: 12px 16px;
          border-bottom: 1px solid #f0f0f0;
          cursor: pointer;
          ${idx === activeIndex ? 'background: #e8f0ff;' : ''}
        `;

        const lz = p.lastZero ?? (p.code && stockIndexGlobal[p.code]?.last_zero_date);
        const codeStyle = lastZeroColor(lz);

        const icon = markerConfig[p.type]?.icon;
        const iconUrl = icon ? icon.options.iconUrl : '';
        const iconHtml = iconUrl ? `<img src="${iconUrl}" style="width: 72px; height: 72px; object-fit: contain;" />` : `<div style="width: 24px; height: 24px;"></div>`;

        li.innerHTML = `
          <div style="display: flex; justify-content: space-between; align-items: flex-start; gap: 12px;">
            <div style="flex: 1;">
              <div style="font-weight: 600; ${codeStyle}">
                ${p.code ?? '—'} — ${p.material}
              </div>
              <div style="color: #666; margin-top: 4px;">
                ${p.name}
              </div>
              ${typeof p.invLbs === 'number' ? `<div style="color:#444;margin-top:3px;font-size:11px">Inventory: <b>${fmtInt(p.invLbs)} lbs</b></div>` : ''}
              ${lz ? `<div style="color:#888;margin-top:3px;font-size:11px">Last Zero: <span style="${codeStyle}">${lz}</span>${p.ageLabel ? ` · <span style="${codeStyle}">${p.ageLabel}</span>` : ''}</div>` : ''}
            </div>
            <div style="display: flex; gap: 6px; align-items: center; flex-shrink: 0;">
              ${iconHtml}
              <button style="
                padding: 4px 8px;
                border: 1px solid #ccc;
                border-radius: 3px;
                background: #f8f8f8;
                cursor: pointer;
                font-size: 11px;
                white-space: nowrap;
              ">Ping</button>
            </div>
          </div>
        `;

        li.addEventListener('click', () => {
          activeIndex = idx;
          renderList();
          pingPileWithStation(p);
          modal.focus();
        });

        const btn = li.querySelector('button');
        btn.addEventListener('click', e => {
          e.stopPropagation();
          pingPileWithStation(p);
          modal.focus();
        });

        pileList.appendChild(li);
      });

      // Restore focus to modal after rendering
      if (document.activeElement !== input) {
        modal.focus();
      }
    }

    // Set mode
    function setMode(mode) {
      searchMode = mode;
      activeIndex = -1; // Reset index when changing mode
      btnAll.style.fontWeight = mode === 'all' ? '700' : '';
      btnPast.style.fontWeight = mode === 'pastDue' ? '700' : '';
      btnExport.style.display = mode === 'pastDue' ? '' : 'none';
      inputContainer.style.display = mode === 'pastDue' ? 'none' : '';
      input.style.display = mode === 'pastDue' ? 'none' : '';
      if (mode === 'all') {
        input.focus();
      }
      if (mode === 'pastDue') {
        input.value = '';
        modal.focus();
      }
      renderList();
    }

    // Event listeners
    btnAll.onclick = () => setMode('all');
    btnPast.onclick = () => setMode('pastDue');
    btnExport.onclick = () => exportPastDueXlsx();
    input.oninput = renderList;

    // Keyboard navigation - attach to modal for consistent behavior in both modes
    function handleKeydown(e) {
      const items = pileList.children;
      if (!items.length) return;

      if (e.key === 'ArrowDown') {
        e.preventDefault();
        e.stopPropagation();
        activeIndex = Math.min(activeIndex + 1, items.length - 1);
        items[activeIndex].scrollIntoView({ block: 'nearest' });
        renderList();
        return;
      }

      if (e.key === 'ArrowUp') {
        e.preventDefault();
        e.stopPropagation();
        activeIndex = Math.max(activeIndex - 1, 0);
        items[activeIndex].scrollIntoView({ block: 'nearest' });
        renderList();
        return;
      }

      if (e.key === 'Enter' && activeIndex >= 0) {
        e.preventDefault();
        e.stopPropagation();
        const p = currentFilteredRows[activeIndex];
        if (p) {
          pingPileWithStation(p);
        }
        return;
      }

      if (e.key === 'Escape') {
        e.preventDefault();
        e.stopPropagation();
        closeSearchModal();
        return;
      }
    }

    input.addEventListener('keydown', handleKeydown);
    modal.addEventListener('keydown', handleKeydown);

    // Initialize
    setMode(searchMode);

    // Close handler
    function closeSearchModal() {
      modal.remove();
      backdrop.remove();
      map.keyboard.enable();
      isSearchModalOpen = false;
      const searchButton = document.getElementById('searchPanelToggle');
      if (searchButton) {
        searchButton.style.display = '';
      }
    }

    closeBtn.onclick = closeSearchModal;

    return closeSearchModal;
  }

  /* ===================================================================
 FLOATING MAGNIFYING GLASS BUTTON
=================================================================== */

  const searchBtnCtrl = L.control({ position: 'bottomright' });
  searchBtnCtrl.onAdd = function () {
    const container = L.DomUtil.create('div');
    L.DomEvent.disableScrollPropagation(container);
    L.DomEvent.disableClickPropagation(container);

    const btn = document.createElement('button');
    btn.id = 'searchPanelToggle';
    btn.title = 'Search Piles';
    btn.style.cssText = `
      width: 40px;
      height: 40px;
      border-radius: 50%;
      background: rgba(255, 255, 255, 0.95);
      border: 2px solid #999;
      cursor: pointer;
      display: flex;
      align-items: center;
      justify-content: center;
      font-size: 20px;
      box-shadow: 0 1px 3px rgba(0, 0, 0, 0.15);
      transition: all 0.2s ease;
      padding: 0;
    `;
    btn.innerHTML = '🔍';

    btn.onmouseover = () => {
      btn.style.background = 'rgba(255, 255, 255, 1)';
      btn.style.boxShadow = '0 2px 5px rgba(0, 0, 0, 0.25)';
    };
    btn.onmouseout = () => {
      btn.style.background = 'rgba(255, 255, 255, 0.95)';
      btn.style.boxShadow = '0 1px 3px rgba(0, 0, 0, 0.15)';
    };

    btn.onclick = () => {
      if (isSearchModalOpen) {
        return;
      }
      map.keyboard.disable();
      createSearchModal();
    };

    container.appendChild(btn);
    return container;
  };

  searchBtnCtrl.addTo(map);

  /* ===================================================================
   CONSUMPTION CSV PARSER
  =================================================================== */
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

 if (unknownTypes.size) console.warn('Unknown types:', Array.from(unknownTypes));
}).catch(err => console.error('Data load failed:', err));

/* ===================================================================
 Past Due XLSX Export (robust, non-corrupt)
=================================================================== */
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

    const rows = pastDue.map(p => {
      const codeKey = normalizePileCode(p.code);
      const invLbs = typeof p.invLbs === 'number' ? Math.max(0, p.invLbs) : 0;
      const avgDaily = pileAvgByCode[codeKey] ?? 0;
      const dud = avgDaily > 0 ? invLbs / avgDaily : 0;

      return {
        'Pile Number': p.code ?? '—',
        'Name': p.name ?? '—',
        'Material': p.material ?? '—',
        'Last Zero Date': p.lastZero ?? '—',
        'Age': p.ageLabel ?? '—',
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
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'PastDue');

    // Filename
    const today = new Date();
    const yyyy = today.getFullYear();
    const mm = String(today.getMonth() + 1).padStart(2, '0');
    const dd = String(today.getDate()).padStart(2, '0');

    XLSX.writeFile(wb, `PastDue_${yyyy}-${mm}-${dd}.xlsx`);

  } catch (err) {
    console.error('Past Due export failed:', err);
    alert('Past Due export failed.');
  }
};

/* ===================================================================
 LAYER CONTROL
=================================================================== */
const overlayMaps = {};
Object.values(markerConfig).forEach(cfg => overlayMaps[cfg.displayName] = cfg.layer);
const layerControl = L.control.layers(null, overlayMaps, {
  collapsed: true,
  position: "bottomleft"
}).addTo(map);

/* ===================================================================
 CHECK ALL/REMOVE ALL BUTTON
=================================================================== */
const checkAllBtn = L.control({ position: "bottomleft" });
checkAllBtn.onAdd = function() {
  const btn = L.DomUtil.create("button");
  btn.textContent = "Check All";
  btn.style.margin = "4px";
  btn.style.padding = "4px 8px";
  btn.style.cursor = "pointer";
  btn.onclick = () => {
    Object.values(markerConfig).forEach(cfg => map.addLayer(cfg.layer));
  };
  return btn;
};
checkAllBtn.addTo(map);

const uncheckAllBtn = L.control({ position: "bottomleft" });
uncheckAllBtn.onAdd = function() {
  const btn = L.DomUtil.create("button");
  btn.textContent = "Remove All";
  btn.style.margin = "4px";
  btn.style.padding = "4px 8px";
  btn.style.cursor = "pointer";
  btn.onclick = () => {
    Object.values(markerConfig).forEach(cfg => { if (map.hasLayer(cfg.layer)) map.removeLayer(cfg.layer); });
    Object.values(loadCells).forEach(lc => { if (map.hasLayer(lc)) map.removeLayer(lc); });
  };
  return btn;
};
uncheckAllBtn.addTo(map);

/* ===================================================================
 CONTEXT MENU
=================================================================== */
const contextMenu = document.createElement('div');
contextMenu.id = 'contextMenu';
contextMenu.innerHTML = `
  <ul>
    <li id="getCoords">📍 Get Coordinates</li>
  </ul>
`;
contextMenu.style.position = 'absolute';
contextMenu.style.display = 'none';
contextMenu.style.zIndex = 2000;
document.body.appendChild(contextMenu);

let clickLatLng = null;
map.on('contextmenu', function(e) {
  clickLatLng = e.latlng;
  contextMenu.style.left = e.originalEvent.pageX + 'px';
  contextMenu.style.top = e.originalEvent.pageY + 'px';
  contextMenu.style.display = 'block';
});
map.on('click', () => contextMenu.style.display = 'none');
document.getElementById('getCoords').addEventListener('click', async () => {
  if (!clickLatLng) return;
  const { lat, lng } = clickLatLng;
  const template = {
    type: "MARKERTYPE",
    name: "PILE# & MATERIAL",
    lat: parseFloat(lat.toFixed(14)),
    lng: parseFloat(lng.toFixed(14))
  };
  await navigator.clipboard.writeText(JSON.stringify(template));
  const emojiIcon = L.divIcon({
    html: '📍',
    className: '',
    iconSize: [48, 48],
    iconAnchor: [8,8]
  });
  const tempMarker = L.marker([lat, lng], { icon: emojiIcon }).addTo(map);
  setTimeout(() => {map.removeLayer(tempMarker);}, 1500);
  contextMenu.style.display = 'none';
});

/* ===================================================================
 GLOBAL ERROR LOGGING
=================================================================== */
window.addEventListener("error", e => console.error("Error:", e.message));

// Pre-load History.xlsx and discover available months so the month selector
// opens instantly without a loading step on the first click.
(async () => {
  try {
    const [rx, bk, burn, brk] = await Promise.all([
      discoverHistoryMonths('Receiving1'),
      discoverHistoryMonths('Consumption1'),
      discoverHistoryMonthsForYearlySheet('BurningHistory'),
      discoverHistoryMonthsForYearlySheet('BreakingHistory')
    ]);
    receivingHistoryMonths = rx;
    bucketHistoryMonths = bk;
    burningHistoryMonths = burn;
    breakingHistoryMonths = brk;
  } catch (err) {
    console.warn('History pre-load failed — month selector will discover on demand:', err);
  }
})();
