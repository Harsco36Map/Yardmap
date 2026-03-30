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
    zoomDelta: 0.25
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
    div.textContent = 'Inventory data current as of —';
    return div;
  };
  ctrl.addTo(map);
  fetchLatestInventoryCsv().then(payload => {
    const d = payload && payload.meta && payload.meta.report_date;
    const banner = document.getElementById('invBanner');
    if (banner && d) {
      // Format date from "DD-MMM-YYYY HH:MM:SS" to "Month-Day" (e.g., "March-26")
      const dateStr = d.trim();
      let formattedDate = dateStr;
      
      // Try to parse DD-MMM-YYYY format (e.g., "25-MAR-2026")
      const dateParts = dateStr.match(/(\d{1,2})-([A-Za-z]{3})-(\d{4})/);
      if (dateParts) {
        const day = parseInt(dateParts[1], 10);
        const monthStr = dateParts[2].toUpperCase();
        const monthMap = {
          'JAN': 'January', 'FEB': 'February', 'MAR': 'March', 'APR': 'April',
          'MAY': 'May', 'JUN': 'June', 'JUL': 'July', 'AUG': 'August',
          'SEP': 'September', 'OCT': 'October', 'NOV': 'November', 'DEC': 'December'
        };
        const monthName = monthMap[monthStr];
        if (monthName) {
          formattedDate = `${monthName}-${day}`;
        }
      }
      // Fallback: also try MM/DD/YYYY format (in case data format changes)
      else {
        const mdy = dateStr.match(/(\d{1,2})\/(\d{1,2})\/(\d{4})/);
        if (mdy) {
          const month = parseInt(mdy[1], 10);
          const day = parseInt(mdy[2], 10);
          const monthNames = ['January', 'February', 'March', 'April', 'May', 'June', 
                            'July', 'August', 'September', 'October', 'November', 'December'];
          if (month >= 1 && month <= 12) {
            formattedDate = `${monthNames[month - 1]}-${day}`;
          }
        }
      }
      
      banner.textContent = `Inventory data current as of ${formattedDate}`;
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

// 2) Cache + helpers --------------------------------------------------
let totalsWorkbookCache = { at: 0, workbook: null };
let burningCache = { at: 0, data: null };
let breakingCache = { at: 0, data: null };
let bucketLoadingCache = { at: 0, data: null };
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

function parseXlsxDateCell(value) {
  if (value instanceof Date && isFinite(value.getTime())) return value;
  const v = String(value || '').trim();
  if (!v) return null;
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

// ===================== GLOBAL HELPERS (must be above builders) =====================

// Leading alphanumerics (e.g., "62U" from "62U Unbreakable")
function extractPileCode(name) {
  if (!name) return null;
  const m = String(name).match(/^[A-Za-z0-9]+/);
  return m ? m[0] : null;
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
    meta: { report_date },
    stock
  };
}

async function fetchBurningTotals(force = false) {
  const now = Date.now();
  if (!force && burningCache.data && (now - burningCache.at) < 120000) {
    return burningCache.data;
  }

  const rows = [];
  const num = v => {
    const x = Number((v || '').toString().trim());
    return Number.isFinite(x) ? x : null;
  };

  const workbook = await loadTotalsWorkbook(force);
  const sheet = findSheetByName(workbook, 'Burning');
  if (!sheet) throw new Error('Burning sheet not found in Production.xlsx');

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

  const { year: filterYear, month: filterMonth } = getCurrentInventoryPeriod();
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
  burningCache = { at: now, data: payload };
  return payload;
}

async function fetchBreakingTotals(force = false) {
  const now = Date.now();
  if (!force && breakingCache.data && (now - breakingCache.at) < 120000) {
    return breakingCache.data;
  }

  const rows = [];
  const toNum = (v) => {
    const x = Number((v || '').toString().trim());
    return Number.isFinite(x) ? x : null;
  };

  const workbook = await loadTotalsWorkbook(force);
  const sheet = findSheetByName(workbook, 'Breaking');
  if (!sheet) throw new Error('Breaking sheet not found in Production.xlsx');

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

  const { year: filterYear, month: filterMonth } = getCurrentInventoryPeriod();
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
  breakingCache = { at: now, data: payload };
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
      <td style="padding:2px 6px;white-space:nowrap">${esc(r.dateLabel)}</td>
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
      &lt;td style="padding:2px 6px;white-space:nowrap"&gt;${esc(r.dateLabel)}&lt;/td&gt;
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
    &lt;div style="font-weight:700;margin-bottom:6px"&gt;Burning Station&lt;/div&gt;
    ${summary}
    ${activity}
    ${coilsSection}
  `;

  return `&lt;div style="${POPUP_CONTAINER_STYLE}"&gt;${body}&lt;/div&gt;`;
}

// 5) Wire popup events ------------------------------------------------

async function fetchBucketLoadingConsumption(force = false) {
  const now = Date.now();
  if (!force && bucketLoadingCache.data && (now - bucketLoadingCache.at) < 120000) {
    return bucketLoadingCache.data;
  }

  const toNum = v => {
    const n = Number(v || 0);
    return Number.isFinite(n) ? n : 0;
  };

  const workbook = await loadTotalsWorkbook(force);
  const sheet = findSheetByName(workbook, 'Consumption');
  if (!sheet) {
    console.warn('Consumption sheet not found');
    return null;
  }

  const data = window.XLSX.utils.sheet_to_json(sheet, { header: 1, raw: false });
  if (!Array.isArray(data) || data.length < 2) {
    console.warn('Consumption sheet has insufficient data');
    return null;
  }

  const { year: filterYear, month: filterMonth } = getCurrentInventoryPeriod();

  const maxCols = data.reduce((m, row) => Math.max(m, Array.isArray(row) ? row.length : 0), 0);
  const candidateStarts = [];
  for (let c = 0; c <= maxCols - 3; c += 3) {
    candidateStarts.push(c);
  }

  let chosenStart = null;
  let bestMatchCount = -1;

  const monthMatches = (date) => {
    return date instanceof Date && !Number.isNaN(date.getTime()) && date.getFullYear() === filterYear && date.getMonth() === filterMonth;
  };

  for (const start of candidateStarts) {
    let matchCount = 0;
    let sampleCount = 0;

    for (let r = 1; r < data.length; r++) {
      const rowData = data[r];
      if (!Array.isArray(rowData)) continue;
      const dateVal = rowData[start];
      if (dateVal == null || String(dateVal).trim() === '') continue;

      sampleCount += 1;
      const parsed = parseXlsxDateCell(dateVal);
      if (monthMatches(parsed)) {
        matchCount += 1;
      }
    }

    if (matchCount > bestMatchCount || (matchCount === bestMatchCount && sampleCount > 0 && chosenStart === null)) {
      bestMatchCount = matchCount;
      chosenStart = start;
    }
  }

  if (chosenStart === null) {
    chosenStart = ((filterMonth + 1) * 3);
    if (chosenStart >= maxCols) chosenStart = 0;
  }

  const rows = [];
  let totalPounds = 0;
  let totalTons = 0;
  let avgPounds = 0;
  let avgTons = 0;
  let totalRowCount = 0;

  for (let r = 1; r < data.length; r++) {
    const rowData = data[r];
    if (!Array.isArray(rowData)) continue;

    const dateCell = rowData[chosenStart];
    const poundsCell = rowData[chosenStart + 1];
    const tonsCell = rowData[chosenStart + 2];

    if ((dateCell == null || String(dateCell).trim() === '') && (poundsCell == null || String(poundsCell).trim() === '') && (tonsCell == null || String(tonsCell).trim() === '')) {
      continue;
    }

    const dateText = String(dateCell || '').trim();
    const lbs = toNum(poundsCell);
    const tons = toNum(tonsCell);
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

    const parsed = parseXlsxDateCell(dateCell);
    if (monthMatches(parsed)) {
      dayNum = parsed.getDate();
      dateLabel = `${parsed.getMonth() + 1}/${parsed.getDate()}`;
    } else {
      const maybeDay = Number(dateText);
      if (Number.isFinite(maybeDay) && maybeDay >= 1 && maybeDay <= 31) {
        dayNum = maybeDay;
        dateLabel = `Day ${maybeDay}`;
      } else {
        continue;
      }
    }

    rows.push({ day: dayNum, dateLabel, pounds: lbs, tons });
    totalRowCount += 1;
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
      blockStart: chosenStart
    },
    totals: {
      totalPounds,
      totalTons,
      avgPounds,
      avgTons
    }
  };

  window.currentBucketLoadingRows = rows;
  bucketLoadingCache = { at: now, data: payload };
  return payload;
}

function renderBucketLoadingPopup(payload) {
  if (!payload) {
    return '<b>Bucket Loading</b><div>No data.</div>';
  }

  const t = payload.totals || {};
  const totalPoundsText = (typeof t.totalPounds === 'number' && isFinite(t.totalPounds)) ? fmtInt(t.totalPounds) : '—';
  const totalTonsText = (typeof t.totalTons === 'number' && isFinite(t.totalTons)) ? fmtTons2(t.totalTons, 2) : '—';
  const avgPoundsText = (typeof t.avgPounds === 'number' && isFinite(t.avgPounds)) ? fmtInt(t.avgPounds) : '—';
  const avgTonsText = (typeof t.avgTons === 'number' && isFinite(t.avgTons)) ? fmtTons2(t.avgTons, 2) : '—';

  const rowsHtml = (payload.rows || []).map(r => `
    <tr>
      <td style="padding:2px 6px;white-space:nowrap">${esc(r.dateLabel)}</td>
      <td style="padding:2px 6px;text-align:right">${fmtInt(r.pounds)}</td>
      <td style="padding:2px 6px;text-align:right">${fmtTons2(r.tons, 2)}</td>
    </tr>
  `).join('');

  const rowCount = (payload.rows || []).length;
  const countText = rowCount > 0 ? `<div style="font-size:11px;color:#999;margin:4px 0;">${rowCount} rows this month</div>` : '';

  const body = `
  <div style="font-weight:700;margin-bottom:6px">Bucket Loading</div>
  <div style="margin-bottom:8px;text-align:center">
    <table style="width:auto;font-size:12px;line-height:1.3;border-collapse:collapse;margin:0 auto;text-align:left">
      <tr><td style="color:#666;padding:2px 6px">Total Consumed</td><td style="text-align:right;padding:2px 6px"><b>${totalPoundsText} lbs</b> <span style="color:#555">(<b>${totalTonsText} tons</b>)</span></td></tr>
      <tr><td style="color:#666;padding:2px 6px">Daily Average</td><td style="text-align:right;padding:2px 6px"><b>${avgPoundsText} lbs</b> <span style="color:#555">(<b>${avgTonsText} tons</b>)</span></td></tr>
    </table>
  </div>
  <div id="bucketLoadingActivity" style="${ACTIVITY_CONTAINER_STYLE}">
    ${countText}
    <table style="${ACTIVITY_TABLE_STYLE}">
      <thead>
        <tr style="background:#f2f2f2">
          <th style="text-align:left;padding:2px 6px">Date</th>
          <th style="text-align:right;padding:2px 6px">Pounds</th>
          <th style="text-align:right;padding:2px 6px">Tons</th>
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

function wireBucketLoadingPopupEvents(container, marker) {
  const toggle = container.querySelector('#bucketLoadingToggle');
  const block = container.querySelector('#bucketLoadingActivity');

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
}

function wireBreakingPopupEvents(container, marker) {
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

burningArea.on('popupopen', async () => {
  burningArea.setPopupContent(`&lt;div style="${POPUP_CONTAINER_STYLE}"&gt;Loading…&lt;/div&gt;`);
  try {
    const payload = await fetchBurningTotals();
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

breakingArea.on('popupopen', async () => {
  breakingArea.setPopupContent(`<div style="${POPUP_CONTAINER_STYLE}">Loading…</div>`);
  try {
    const payload = await fetchBreakingTotals();
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

bucketLoadingArea.on('popupopen', async () => {
  bucketLoadingArea.setPopupContent(`<div style="${POPUP_CONTAINER_STYLE}">Loading…</div>`);
  try {
    const payload = await fetchBucketLoadingConsumption();
    const encoded = renderBucketLoadingPopup(payload);
    const decoded = unescapeAngles(encoded);
    bucketLoadingArea.setPopupContent(decoded);

    setTimeout(() => {
      const el = bucketLoadingArea.getPopup()?.getElement();
      if (el) wireBucketLoadingPopupEvents(el, bucketLoadingArea);
    }, 0);
  } catch (err) {
    console.error(err);
    bucketLoadingArea.setPopupContent('<b>Bucket Loading</b><div style="color:#c00">Failed to load Consumption sheet.</div>');
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
      return {
        code,
        name: m.name,
        material: s?.material ?? '',
        marker: m._leaflet,
        type: m.type
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
    // A: Day | B: Consumed | C: Net_Tons | D: blank | E: Pile | F: Total_Actual | G: Avg_Daily
    const lines = text.split(/\r?\n/);
    if (!lines.length) throw new Error("Consumption CSV is empty.");
    const header = (lines[0] || "").trim();
    const expected = "Day,Consumed,Net_Tons,,Pile,Total_Actual,Avg_Daily";
    if (!header.startsWith(expected)) {
      console.warn("Consumption CSV header did not match side-by-side format:", header);
    }
    const pileAvgByCode = {}; // { CODE -> avgDaily }
    for (let i = 1; i < lines.length; i++) {
      const raw = lines[i];
      if (!raw) continue;
      const parts = raw.split(",");
      if (parts.length < 7) continue;
      const pile = (parts[4] || "").trim().toUpperCase(); // column E
      const avgStr = (parts[6] || "").trim();             // column G
      if (!pile) continue;
      const avg = Number(avgStr);
      pileAvgByCode[pile] = Number.isFinite(avg) ? avg : 0;
    }
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
``

    // Fetch consumption averages
    const { pileAvgByCode } = await fetchConsumptionCsv();

    const rows = pastDue.map(p => {
      const codeKey = p.code?.trim().toUpperCase() ?? '';
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
