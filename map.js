
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
  fetch('stockData.json').then(r => r.json()).then(payload => {
    const d = payload && payload.meta && payload.meta.report_date;
    const banner = document.getElementById('invBanner');
    if (banner && d) banner.textContent = `Inventory data current as of ${d}`;
  }).catch(err => console.warn('stockData.json meta fetch failed:', err));
})();


/* ===================================================================
   BURNING STATION  — stand-alone (not in markers.json / overlay)
   =================================================================== */

// 1) Config -----------------------------------------------------------
const burningCsvUrl = 'BurningTotals.csv';
const burningLatLng = [40.79365495632949, -82.53501357377355];
const breakingCsvUrl = 'BreakingTotals.csv';
const breakingLatLng = [40.79408763021845, -82.5388475264302];

// 2) Cache + helpers --------------------------------------------------
let burningCache = { at: 0, data: null };
let breakingCache = { at: 0, data: null };

function fmtInt(n)  { return (typeof n === 'number' && isFinite(n)) ? Math.round(n).toLocaleString('en-US') : '—'; }
function fmtTons(n, d = 3) { return (typeof n === 'number' && isFinite(n)) ? n.toFixed(d) : '—'; }
function fmtTons2(n, d = 2) { return (typeof n === 'number' && isFinite(n)) ? n.toFixed(d) : '—'; }

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
async function fetchBurningTotals(force = false) {
  const now = Date.now();
  if (!force && burningCache.data && (now - burningCache.at) < 120000) {
    return burningCache.data;
  }

  const res = await fetch(burningCsvUrl, { cache: 'no-store' });
  if (!res.ok) throw new Error(`Failed to fetch ${burningCsvUrl}: ${res.status}`);
  const text = await res.text();

  const lines = text.split(/\r?\n/).filter(Boolean);
  if (!lines.length) throw new Error('BurningTotals.csv is empty.');

  const rows = [];
  for (let i = 1; i < lines.length; i++) {
    const p = lines[i].split(',');
    if (p.length < 7) continue;

    const dateStr = (p[0] || '').trim();
    if (!dateStr) continue;

    const dt = new Date(dateStr);
    if (!isFinite(dt.getTime())) continue;

    const num = v => {
      const x = Number((v || '').trim());
      return Number.isFinite(x) ? x : null;
    };

    rows.push({
      date: dt,
      dateLabel: dateStr,
      netLbs: num(p[1]),
      from: (p[2] || '').trim(),
      to: (p[3] || '').trim(),
      netTons: num(p[4]),
      cuts: num(p[5]),
      billableTons: num(p[6])
    });
  }

  rows.sort((a,b)=>b.date-a.date);

  let totalLbs = 0, totalTons = 0, totalCuts = 0, totalBillable = 0, latest=null;
  for (const r of rows) {
    totalLbs      += r.netLbs || 0;
    totalTons     += r.netTons || 0;
    totalCuts     += r.cuts   || 0;
    totalBillable += r.billableTons || 0;
    if (!latest || r.date > latest) latest = r.date;
  }

  const payload = {
    rows,
    totals: { netLbs: totalLbs, netTons: totalTons, cuts: totalCuts, billableTons: totalBillable },
    latestDateLabel: latest ? latest.toISOString().slice(0,10) : '—'
  };

  burningCache = { at: now, data: payload };
  return payload;
}

async function fetchBreakingTotals(force = false) {
  const now = Date.now();
  if (!force && breakingCache.data && (now - breakingCache.at) < 120000) {
    return breakingCache.data;
  }
  const res = await fetch(breakingCsvUrl, { cache: 'no-store' });
  if (!res.ok) throw new Error(`Failed to fetch ${breakingCsvUrl}: ${res.status}`);

  const text = await res.text();
  const lines = text.split(/\r?\n/).filter(Boolean);
  if (!lines.length) throw new Error('BreakingTotals.csv is empty.');

  const rows = [];
  for (let i = 1; i < lines.length; i++) {
    const p = lines[i].split(',');
    if (p.length < 5) continue;

    const dateStr = (p[0] || '').trim();
    if (!dateStr) continue;
    const dt = new Date(dateStr);
    if (!isFinite(dt.getTime())) continue;

    const num = v => {
      const x = Number((v || '').trim());
      return Number.isFinite(x) ? x : null;
    };

    const from = (p[2] || '').trim();
    const to = (p[3] || '').trim();
    const netTons = num(p[4]);
    const material = (p[7] || '').trim(); // optional; use if present

    rows.push({ date: dt, dateLabel: dateStr, from, to, netTons, material });
  }

  rows.sort((a,b) => b.date - a.date);

  // Summaries
  const nowD = new Date();
  const Y = nowD.getFullYear(), M = nowD.getMonth();
  const monthRows = rows.filter(r => r.date.getFullYear() === Y && r.date.getMonth() === M);
  const totalMonthTons = monthRows.reduce((a, r) => a + (r.netTons || 0), 0);
  const materialsTouched = Array.from(new Set(
    monthRows.map(r => (r.material || '').trim()).filter(Boolean)
  )).sort();

  const payload = {
    rows,
    month: {
      year: Y,
      monthIndex: M, // 0..11
      totalMonthTons,
      materialsTouched,
      rowCount: monthRows.length
    }
  };
  breakingCache = { at: now, data: payload };
  return payload;
}

// 4) Popup rendering

  function extractPileCode(name) {
    if (!name) return null;
    const m = name.match(/^[A-Za-z0-9]+/); // leading alphanumerics (e.g., "62U" from "62U Unbreakable")
    return m ? m[0] : null;
  }

function buildUnprepRows(markers, stockIndex) {
  return markers
    .filter(m => m.type === "Breaking" || m.type === "Unbreakable")
    .map(m => {
      const code = extractPileCode(m.name);
      const s = (code && stockIndex[code]) ? stockIndex[code] : {};
      const inv = s.operating_inventory_lbs ?? 0;
      const lastZero = s.last_zero_date ?? '—';
      return `
        <tr>
          <td style="padding:2px 6px">${m.name}</td>
          <td style="padding:2px 6px;text-align:right">${inv.toLocaleString('en-US')}</td>
          <td style="padding:2px 6px">${lastZero}</td>
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
      const stock = stockIndex[code] || {};
      const inv = stock.operating_inventory_lbs ?? 0;
      const lastZero = stock.last_zero_date ?? '—';

      return `
        <tr>
          <td style="padding:2px 6px">${m.name}</td>
          <td style="padding:2px 6px;text-align:right">${inv.toLocaleString()}</td>
          <td style="padding:2px 6px">${lastZero}</td>
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
``

function renderBreakingPopup(payload, markers, stockIndex) {
  if (!payload) {
    return '<b>Breaking Pit</b><div>No data.</div>';
  }

  // Monthly summary pieces
  const monthTons = payload.month.totalMonthTons;
  const monthTonsText = fmtTons2(monthTons, 2);
  const materials = payload.month.materialsTouched;
  const materialsText = materials.length ? materials.join(', ') : '—';
  const unprepTotalLbs = getUnprepTotalInventory(markers, stockIndex);
  const unprepLbsText  = isFinite(unprepTotalLbs) ? unprepTotalLbs.toLocaleString('en-US') : '—';
  const unprepTonsText = isFinite(unprepTotalLbs) ? (unprepTotalLbs / 2000).toFixed(2) : '—';

  // Summary block
  const summary = `
    <div style="margin-bottom:8px">
      <table style="width:100%;font-size:12px;line-height:1.3;border-collapse:collapse">
        <tr>
          <td style="color:#666;padding:2px 6px">Total Unprep Inventory</td>
          <td style="text-align:right;padding:2px 6px"><b>${unprepLbsText}</b> <span style="color:#555">(${unprepTonsText} tons)</span></td>
        </tr>
        <tr>
          <td style="color:#666;padding:2px 6px">Total Processed(net tons)</td>
          <td style="text-align:right;padding:2px 6px"><b>${monthTonsText}</b></td>
        </tr>
      </table>
    </div>
  `;

  // Activity rows (you can choose monthRows only; here we show ALL rows like Burning)
  const rowsHtml = payload.rows.map(r => `
    <tr>
      <td style="padding:2px 6px;white-space:nowrap">${esc(r.dateLabel)}</td>
      <td style="padding:2px 6px">${esc(r.from)} → ${esc(r.to)}</td>
      <td style="padding:2px 6px;text-align:right">${fmtTons2(r.netTons, 3)}</td>
      <td style="padding:2px 6px">${esc(r.material || '')}</td>
    </tr>
  `).join('');

  const activity = `
  <div id="breakingActivity" style="display:none">
    <table style="width:100%;font-size:12px;border-collapse:collapse">
      <thead>
        <tr style="background:#f2f2f2">
          <th style="text-align:left;padding:2px 6px">Date</th>
          <th style="text-align:left;padding:2px 6px">From → To</th>
          <th style="text-align:right;padding:2px 6px">NetTons</th>
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

    <!-- Top Unprep toggle (like CoilsTop) -->
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
  return `<div style="min-width:300px">${body}</div>`;
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
  <div style="margin-bottom:8px">
    <table style="width:100%;font-size:12px;line-height:1.3;border-collapse:collapse">
      <tr>
        <td style="color:#666;padding:2px 6px">Total Coil Inventory</td>
        <td style="text-align:right;padding:2px 6px"><b>${coilsInvText}</b> <span style="color:#555">(${coilsInvTonsText} tons)</span></td>
      </tr>
      <tr>
        <td style="color:#666;padding:2px 6px">Net Tons Cut</td>
        <td style="text-align:right;padding:2px 6px"><b>${fmtTons(t.netTons)}</b></td>
      </tr>
      <tr>
        <td style="color:#666;padding:2px 6px">Billable Tons Cut</td>
        <td style="text-align:right;padding:2px 6px"><b>${fmtTons(t.billableTons)}</b></td>
      </tr>
    </table>
  </div>
`;

  // ALL activity rows
  const rowsHtml = payload.rows.map(r => `
    &lt;tr&gt;
      &lt;td style="padding:2px 6px;white-space:nowrap"&gt;${esc(r.dateLabel)}&lt;/td&gt;
      &lt;td style="padding:2px 6px"&gt;${esc(r.from)} → ${esc(r.to)}&lt;/td&gt;
      &lt;td style="padding:2px 6px;text-align:right"&gt;${fmtTons(r.netTons)}&lt;/td&gt;
      &lt;td style="padding:2px 6px;text-align:right"&gt;${fmtInt(r.cuts)}&lt;/td&gt;
      &lt;td style="padding:2px 6px;text-align:right"&gt;${fmtTons(r.billableTons)}&lt;/td&gt;
    &lt;/tr&gt;
  `).join('');

  // Activity section
  const activity = `
  &lt;div id="burningActivity" style="display:none"&gt;
      &lt;table style="width:100%;font-size:12px;border-collapse:collapse"&gt;
        &lt;thead&gt;
          &lt;tr style="background:#f2f2f2"&gt;
            &lt;th style="text-align:left;padding:2px 6px"&gt;Date&lt;/th&gt;
            &lt;th style="text-align:left;padding:2px 6px"&gt;From → To&lt;/th&gt;
            &lt;th style="text-align:right;padding:2px 6px"&gt;NetTons&lt;/th&gt;
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

  return `&lt;div style="min-width:300px"&gt;${body}&lt;/div&gt;`;
}

// 5) Wire popup events ------------------------------------------------

function wireBreakingPopupEvents(container, marker) {
  // Activity
  const activityToggle = container.querySelector('#breakingToggle');
  const activityBlock  = container.querySelector('#breakingActivity');

  // Unprep top/bottom buttons + section
  const unprepTopBtn = container.querySelector('#unprepToggleTop');
  const unprepBtmBtn = container.querySelector('#unprepToggleBottom');
  const unprepDiv    = container.querySelector('#unprepSection');

  // ---- helpers ----
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
    // hide Activity button while Unprep is open (mirroring Burning)
    if (activityToggle) activityToggle.style.display = show ? 'none' : '';
  }

  // Initialize unprep (hidden by default)
  if (unprepDiv && (unprepTopBtn || unprepBtmBtn)) {
    setUnprepState(unprepDiv.style.display === 'block');
    const onUnprepClick = () => {
      const showing = unprepDiv.style.display === 'block';
      setUnprepState(!showing);
    };
    if (unprepTopBtn) unprepTopBtn.addEventListener('click', onUnprepClick);
    if (unprepBtmBtn) unprepBtmBtn.addEventListener('click', onUnprepClick);
  }

  // Activity toggle (mirrors your Burning Activity wiring)
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
        // only re-show top unprep button if unprep section is closed
        if (unprepTopBtn && (!unprepDiv || unprepDiv.style.display !== 'block')) {
          unprepTopBtn.style.display = '';
        }
      }
    });
    activityBlock.addEventListener('click', e => e.stopPropagation());
  }
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
      const a = document.createElement('a');
      a.href = burningCsvUrl;
      a.target = '_blank';
      a.rel = 'noopener';
      document.body.appendChild(a);
      a.click();
      document.body.removeChild(a);
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
  burningArea.setPopupContent('&lt;div style="min-width:300px"&gt;Loading…&lt;/div&gt;');
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
    burningArea.setPopupContent('&lt;b&gt;Burning Station&lt;/b&gt;&lt;div style="color:#c00"&gt;Failed to load BurningTotals.csv.&lt;/div&gt;');
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
  breakingArea.setPopupContent('<div style="min-width:300px">Loading…</div>');
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
    breakingArea.setPopupContent('<b>Breaking Pit</b><div style="color:#c00">Failed to load BreakingTotals.csv.</div>');
  }
});
``
/* ===================================================================
 LOAD MARKERS + ENRICH POPUPS
=================================================================== */
Promise.all([
  fetch('markers.json').then(r => r.json()),
  fetch('stockData.json').then(r => r.json())
]).then(([markers, stockPayload]) => {
  allMarkersData = markers;                 // save globally
  stockIndexGlobal = stockPayload.stock || {};
  
  const unknownTypes = new Set();

  // Helpers
  function parseMDY(dateStr) {
    if (!dateStr) return null;
    const parts = dateStr.split('/');
    if (parts.length !== 3) return null;
    const m = parseInt(parts[0], 10);
    const d = parseInt(parts[1], 10);
    const y = parseInt(parts[2], 10);
    if (!m || !d || !y) return null;
    return new Date(y, m - 1, d);
  }
  function monthsDiff(a, b) {
    const years = b.getFullYear() - a.getFullYear();
    const months = years * 12 + (b.getMonth() - a.getMonth());
    return (b.getDate() >= a.getDate()) ? months : months - 1;
  }
  function lastZeroColor(dateStr) {
    const dt = parseMDY(dateStr);
    if (!dt) return '';
    const now = new Date();
    const m = monthsDiff(dt, now);
    if (m > 6) return 'color:#c62828; font-weight:700;';
    if (m > 3) return 'color:#f9a825; font-weight:700;';
    return '';
  }
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
    const invText = (typeof s.operating_inventory_lbs === 'number')
      ? s.operating_inventory_lbs.toLocaleString('en-US', { maximumFractionDigits: 0 }) + ' lbs'
      : '—';
    const lzStyle = lastZeroColor(s.last_zero_date);
    return `
      <div style="min-width:220px">
        <div style="font-weight:700;margin-bottom:4px">${marker.name}</div>
        <table style="font-size:12px;line-height:1.3">
          <tr><td style="padding-right:8px;color:#666">Material:</td><td>${s.material || ''}</td></tr>
          <tr><td style="padding-right:8px;color:#666">Inventory:</td><td>${invText}</td></tr>
          <tr><td style="padding-right:8px;color:#666">Last Zero Date:</td><td style="${lzStyle}">${s.last_zero_date || ''}</td></tr>
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

  /* ===================================================================
   Toggleable Past Due Panel
  =================================================================== */
  const pastDueCtrl = L.control({ position: 'bottomright' });
  pastDueCtrl.onAdd = function () {
    const div = L.DomUtil.create('div'); 
    L.DomEvent.disableScrollPropagation(div);
    L.DomEvent.disableClickPropagation(div);
    div.id = 'pastDuePanel';
    Object.assign(div.style, {
      background: 'rgba(255,255,255,0.95)',
      border: '1px solid #ccc',
      borderRadius: '4px',
      padding: '6px 10px',
      margin: '8px',
      font: '12px/1.3 system-ui, -apple-system, Segoe UI, Roboto, Arial, sans-serif',
      boxShadow: '0 1px 3px rgba(0,0,0,0.15)',
      width: '260px',
      maxHeight: '300px',
      overflow: 'hidden',
      transition: 'height 160ms ease, padding 160ms ease'
    });
    let collapsed = true;
    const headerEl = document.createElement('div');
    const bodyEl = document.createElement('div');

    
   // expose simple API to the accordion
   function isExpanded() { return !collapsed; }
   function expand()   { if (collapsed) { collapsed = false; renderHeader(); renderBody(); } }
   function collapse() { if (!collapsed) { collapsed = true;  renderHeader(); renderBody(); } }


    function renderHeader() {
      headerEl.innerHTML = `
        <div style="display:flex;align-items:center;justify-content:space-between">
          <span style="font-weight:700">Past Due (> 6 mo)
            <span style="color:#c62828">●</span> (${pastDue.length})
          </span>
          <button id="pdToggleBtn"
            aria-expanded="${!collapsed}"
            title="${collapsed ? 'Expand' : 'Collapse'}"
            style="padding:2px 6px;border:1px solid #ccc;border-radius:3px;
                   cursor:default;background:#f8f8f8;display:flex;align-items:center">
            ${collapsed ? '▸' : '▾'}
          </button>
        </div>
      `;
    }
    function renderBody() {
      if (collapsed) {
        bodyEl.innerHTML = '';
        div.style.height = '24px';
        div.style.padding = '6px 10px';
        div.style.pointerEvents = 'auto';
        return;
      }
      bodyEl.innerHTML = `
        <div style="margin-top:6px;display:flex;align-items:center;justify-content:space-between">
          <input id="pdSearch" type="text" placeholder="Filter by code/material..."
            style="flex:1;padding:4px 6px;border:1px solid #ddd;border-radius:3px;margin-right:6px">
          <button id="pdExportBtn" title="Export XLSX"
            style="padding:4px 8px;border:1px solid #ccc;border-radius:3px;cursor:pointer;background:#f8f8f8">
            Export
          </button>
        </div>
        <ul id="pdList" style="list-style:none;padding:0;margin:8px 0 0 0;max-height:220px;overflow:auto"></ul>
      `;
      div.style.height = '300px';
      div.style.padding = '6px 10px';

      const ul = bodyEl.querySelector('#pdList');
      const searchInput = bodyEl.querySelector('#pdSearch');

      function renderList(filterText = '') {
        ul.innerHTML = '';
        const term = filterText.trim().toLowerCase();
        const items = pastDue.filter(p => {
          if (!term) return true;
          return (
            (p.code && p.code.toLowerCase().includes(term)) ||
            (p.name && p.name.toLowerCase().includes(term)) ||
            (p.material && p.material.toLowerCase().includes(term))
          );
        });
        items.forEach(p => {
          const li = document.createElement('li');
          li.style.padding = '6px 0';
          li.style.borderBottom = '1px dashed #eee';
          li.style.cursor = 'default';
          const invText = (typeof p.invLbs === 'number')
            ? p.invLbs.toLocaleString('en-US', { maximumFractionDigits: 0 }) + ' lbs'
            : '—';
          li.innerHTML = `
            <div style="display:flex;justify-content:space-between;align-items:center">
              <div>
                <div style="font-weight:600">${p.code} — ${p.material}</div>
                <div style="color:#555">${p.name}</div>
                <div style="color:#c62828;font-weight:700">Last zero: ${p.lastZero}</div>
                <div style="color:#333">Age: ${p.ageLabel}</div>
                <div style="color:#333">Inventory: ${invText}</div>
              </div>
              <button style="margin-left:8px;padding:2px 6px;border:1px solid #ccc;border-radius:3px;cursor:pointer">Ping</button>
              </div>
          `;          
          li.addEventListener('click', () => {
            const target =
              (p.rawType === 'Coils' && window.burningArea) ? window.burningArea :
              ((p.rawType === 'Breaking' || p.rawType === 'Unbreakable') && window.breakingArea) ? window.breakingArea :
              p.marker;
            if (target) pingMarker(target);
          });          
          li.querySelector('button').addEventListener('click', (e) => {
            e.stopPropagation();
            const target =
              (p.rawType === 'Coils' && window.burningArea) ? window.burningArea :
              ((p.rawType === 'Breaking' || p.rawType === 'Unbreakable') && window.breakingArea) ? window.breakingArea :
              p.marker;
            if (target) pingMarker(target);
          });
          ul.appendChild(li);
        });
      }
      renderList();
      searchInput.addEventListener('input', (e) => renderList(e.target.value));

/* ===================== XLSX EXPORT ===================== */
bodyEl.querySelector('#pdExportBtn').addEventListener('click', async () => {
  try {
    if (typeof JSZip === 'undefined') {
      alert('JSZip is required to export XLSX. Please include JSZip in the page.');
      return;
    }

    // Read Avg_Daily from averageconsumption.csv (column G)
    const { pileAvgByCode } = await fetchConsumptionCsv();

    // Timestamp for filename
    const today = new Date();
    const yyyy = today.getFullYear();
    const mm = String(today.getMonth() + 1).padStart(2, '0');
    const dd = String(today.getDate()).padStart(2, '0');

    // Header row
    const header = [
      'Pile Number', 'Name', 'Material', 'Last Zero Date', 'Age', 'Inventory (lbs)',
      'Average Consumed Daily', 'Days Until Depleted', 'Action'
    ];

    // Build dataRows (A..I only)
    const dataRows = pastDue.map(p => {
      const codeKey = p.code ? p.code.trim().toUpperCase() : '';
      const invLbs = (typeof p.invLbs === 'number') ? Math.max(0, p.invLbs) : 0;
      const pileAvgDaily = pileAvgByCode[codeKey] ?? 0;
      const dudPile = (pileAvgDaily > 0) ? (invLbs / pileAvgDaily) : 0;

      return [
        p.code ?? '—',
        p.name ?? '—',
        p.material ?? '—',
        p.lastZero ?? '—',
        p.ageLabel ?? '—',
        Number.isFinite(invLbs) ? invLbs : 0,
        (pileAvgDaily > 0 && isFinite(pileAvgDaily)) ? Math.round(pileAvgDaily) : 0,
        (pileAvgDaily > 0 && isFinite(dudPile)) ? Number(dudPile.toFixed(1)) : 0,
        '' // Action (I)
      ];
    });

    // ==== Auto column widths (Calibri 11 heuristic: pixels ≈ 7*chars + 5) ====
    function computeAutoColWidths(header, dataRows) {
      const cols = header.length; // 9 (A..I)
      const maxChars = Array(cols).fill(0);
      const measure = (val) => {
        if (val === null || val === undefined) return 0;
        const s = (typeof val === 'number') ? String(val) : String(val);
        return s.length;
      };
      for (let c = 0; c < cols; c++) maxChars[c] = Math.max(maxChars[c], measure(header[c]));
      for (const row of dataRows) {
        for (let c = 0; c < cols; c++) maxChars[c] = Math.max(maxChars[c], measure(row[c]));
      }
      const maxDigitWidth = 7; // px per "0" in Calibri 11
      const paddingPx = 5;
      const minWidthCh = 8.43; // don't go narrower than Excel default
      return maxChars.map(chars => {
        const pixels = chars * maxDigitWidth + paddingPx;
        const width = Math.trunc((pixels / maxDigitWidth) * 256) / 256;
        return Math.max(width, minWidthCh);
      });
    }
    const autoColWidths = computeAutoColWidths(header, dataRows);

    // ==== Build XLSX ====
    const zip = new JSZip();
    const xlNS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';

    // Simple XML escaper for inline string cells
    const xmlEsc = s => String(s)
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;')
      .replace(/'/g, '&apos;');

    function colLetter(n) {
      let s = '';
      while (n > 0) {
        const m = (n - 1) % 26;
        s = String.fromCharCode(65 + m) + s;
        n = Math.floor((n - 1) / 26);
      }
      return s;
    }
    function cellRef(r, c) { return `${colLetter(c)}${r}`; }

    const rowCount = dataRows.length + 1;

    function buildCell(r, c, v, styleId = null) {
      const ref = cellRef(r, c);
      if (v === '' || v === null || v === undefined)
        return `<c r="${ref}"${styleId !== null ? ` s="${styleId}"` : ''}/>`;
      if (typeof v === 'number')
        return `<c r="${ref}"${styleId !== null ? ` s="${styleId}"` : ''}><v>${v}</v></c>`;
      return `<c r="${ref}" t="inlineStr"${styleId !== null ? ` s="${styleId}"` : ''}><is><t>${xmlEsc(v)}</t></is></c>`;
    }

    // Sheet rows (header style s=1, numeric right-align s=2 for F..H)
    let sheetData = `<row r="1">`;
    header.forEach((h, i) => { sheetData += buildCell(1, i + 1, h, 1); });
    sheetData += `</row>`;
    dataRows.forEach((row, idx) => {
      const r = idx + 2;
      sheetData += `<row r="${r}">`;
      row.forEach((v, j) => {
        const styleId = (j >= 5 && j <= 7) ? 2 : 0; // F(6),G(7),H(8) right-aligned
        sheetData += buildCell(r, j + 1, v, styleId);
      });
      sheetData += `</row>`;
    });

    // Dimension and <cols> (A..I only; no helper column)
    const firstRef = 'A1';
    const lastRef  = cellRef(rowCount, 9); // I
    let colsXml = '<cols>';
    autoColWidths.forEach((w, i) => {
      const idx = i + 1; // A=1..I=9
      colsXml += `<col min="${idx}" max="${idx}" width="${w.toFixed(2)}" bestFit="1" customWidth="1"/>`;
    });
    colsXml += '</cols>';

    // CF: color entire row A..I based on Action in I (no helper col)
    const dvRange = `I2:I${rowCount}`;
    const cfRange = `$A$2:$I$${rowCount}`;

    const sheet1Xml = `<?xml version="1.0" encoding="UTF-8"?>
<worksheet xmlns="${xlNS}">
  <sheetPr/>
  <dimension ref="${firstRef}:${lastRef}"/>
  <sheetViews><sheetView workbookViewId="0"/></sheetViews>
  <sheetFormatPr defaultRowHeight="15"/>
  ${colsXml}
  <sheetData>
${sheetData}
  </sheetData>

  <conditionalFormatting sqref="${cfRange}">
    <cfRule type="expression" priority="1" dxfId="0"><formula>$I2="Depleted"</formula></cfRule>
    <cfRule type="expression" priority="2" dxfId="1"><formula>$I2="Priority Target"</formula></cfRule>
    <cfRule type="expression" priority="3" dxfId="2"><formula>$I2="Stop Receiving"</formula></cfRule>
  </conditionalFormatting>

  <dataValidations count="1">
    <dataValidation type="list" allowBlank="1" sqref="${dvRange}">
      <formula1>"Depleted,Priority Target,Stop Receiving"</formula1>
    </dataValidation>
  </dataValidations>

  <pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>
</worksheet>`;

    // Package parts
    const contentTypesXml =
      `<?xml version="1.0" encoding="UTF-8"?>` +
      `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">` +
      `<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>` +
      `<Default Extension="xml" ContentType="application/xml"/>` +
      `<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>` +
      `<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>` +
      `<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>` +
      `</Types>`;

    const rootRelsXml =
      `<?xml version="1.0" encoding="UTF-8"?>` +
      `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
      `<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>` +
      `</Relationships>`;

    const workbookXml =
      `<?xml version="1.0" encoding="UTF-8"?>` +
      `<workbook xmlns="${xlNS}" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">` +
      `<sheets><sheet name="PastDue" sheetId="1" r:id="rId1"/></sheets>` +
      `<calcPr calcMode="auto" fullCalcOnLoad="1"/>` +
      `</workbook>`;

    const wbRelsXml =
      `<?xml version="1.0" encoding="UTF-8"?>` +
      `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
      `<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>` +
      `<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>` +
      `</Relationships>`;

    // styles.xml — keep your header (fontId=1), numeric right align (s=2),
    // and add 3 DXFs with BOTH fgColor + bgColor so fills render on Desktop.
    const stylesXml = `<?xml version="1.0" encoding="UTF-8"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <fonts count="2">
    <font><sz val="11"/><name val="Calibri"/></font>
    <font><b/><sz val="11"/><name val="Calibri"/></font>
  </fonts>
  <fills count="2">
    <fill><patternFill patternType="none"/></fill>
    <fill><patternFill patternType="gray125"/></fill>
  </fills>
  <borders count="1">
    <border><left/><right/><top/><bottom/><diagonal/></border>
  </borders>
  <cellStyleXfs count="1">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>
  </cellStyleXfs>
  <cellXfs count="3">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
    <xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0"/> <!-- header bold -->
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyAlignment="1"><alignment horizontal="right"/></xf>
  </cellXfs>
  <cellStyles count="1">
    <cellStyle name="Normal" xfId="0" builtinId="0"/>
  </cellStyles>
  <dxfs count="3">
    <!-- 0: Depleted (light gray) -->
    <dxf>
      <font><color rgb="FF000000"/></font>
      <fill><patternFill patternType="solid">
        <fgColor rgb="FFDDDDDD"/><bgColor rgb="FFDDDDDD"/>
      </patternFill></fill>
    </dxf>
    <!-- 1: Priority Target (light yellow) -->
    <dxf>
      <font><color rgb="FF000000"/></font>
      <fill><patternFill patternType="solid">
        <fgColor rgb="FFFFEB9C"/><bgColor rgb="FFFFEB9C"/>
      </patternFill></fill>
    </dxf>
    <!-- 2: Stop Receiving (light red) -->
    <dxf>
      <font><color rgb="FF000000"/></font>
      <fill><patternFill patternType="solid">
        <fgColor rgb="FFFFC7CE"/><bgColor rgb="FFFFC7CE"/>
      </patternFill></fill>
    </dxf>
  </dxfs>
</styleSheet>`;

    // Add parts to zip
    zip.file('[Content_Types].xml', contentTypesXml);
    zip.folder('_rels').file('.rels', rootRelsXml);
    const xl = zip.folder('xl');
    xl.file('workbook.xml', workbookXml);
    xl.folder('_rels').file('workbook.xml.rels', wbRelsXml);
    xl.folder('worksheets').file('sheet1.xml', sheet1Xml);
    xl.file('styles.xml', stylesXml);

    // Generate & download
    const blob = await zip.generateAsync({ type: 'blob' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = `PastDue_${yyyy}-${mm}-${dd}.xlsx`;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
  } catch (err) {
    console.error('XLSX export failed:', err);
    alert('Export failed: ' + err.message);
  }
});
/* =================== END XLSX EXPORT (auto-width + fixed DXFs + no helper) =================== */

    }

    renderHeader();
    renderBody();
    headerEl.addEventListener('click', (e) => {
      const btn = headerEl.querySelector('#pdToggleBtn');
      if (btn && e.target === btn) { /* keep default */ }
     const wasCollapsed = collapsed;
     collapsed = !collapsed;
     renderHeader();
     renderBody();
     // If we just expanded, collapse the other panel
     if (!collapsed) collapseOthers('pastDuePanel');
    });
    div.appendChild(headerEl);
    div.appendChild(bodyEl);
     registerPanel('pastDuePanel', { isExpanded, expand, collapse });
    return div;
  };

  /* ===================================================================
   Search Panel
  =================================================================== */
  const searchCtrl = L.control({ position: 'bottomright' });
  searchCtrl.onAdd = function () {
    const div = L.DomUtil.create('div');
    L.DomEvent.disableScrollPropagation(div);
    L.DomEvent.disableClickPropagation(div);
    div.id = 'searchPanel';
    Object.assign(div.style, {
      background: 'rgba(255,255,255,0.95)',
      border: '1px solid #ccc',
      borderRadius: '4px',
      padding: '6px 10px',
      margin: '8px',
      font: '12px/1.3 system-ui, -apple-system, Segoe UI, Roboto, Arial, sans-serif',
      boxShadow: '0 1px 3px rgba(0,0,0,0.15)',
      width: '260px',
      maxHeight: '300px',
      overflow: 'hidden',
      transition: 'height 160ms ease, padding 160ms ease'
    });
    let collapsed = true;
    const headerEl = document.createElement('div');
    const bodyEl = document.createElement('div');    
   // expose simple API to the accordion
   function isExpanded() { return !collapsed; }
   function expand()   { if (collapsed) { collapsed = false; renderHeader(); renderBody(); } }
   function collapse() { if (!collapsed) { collapsed = true;  renderHeader(); renderBody(); } }
    function renderHeader() {
      headerEl.innerHTML = `
        <div style="display:flex;align-items:center;justify-content:space-between">
          <span style="font-weight:700">Search Piles</span>
          <button id="srchToggleBtn"
            aria-expanded="${!collapsed}"
            title="${collapsed ? 'Expand' : 'Collapse'}"
            style="padding:2px 6px;border:1px solid #ccc;border-radius:3px;cursor:default;background:#f8f8f8">
            ${collapsed ? '▸' : '▾'}
          </button>
        </div>
      `;
    }
    function renderBody() {
      if (collapsed) {
        bodyEl.innerHTML = '';
        div.style.height = '24px';
        div.style.padding = '6px 10px';
        div.style.pointerEvents = 'auto';
        return;
      }
      bodyEl.innerHTML = `
        <div style="margin-top:6px;display:flex;align-items:center;justify-content:space-between">
          <input id="pileSearch" type="text" placeholder="Filter by code/name/material..."
            style="flex:1;padding:4px 6px;border:1px solid #ddd;border-radius:3px">
        </div>
        <ul id="searchList" style="list-style:none;padding:0;margin:8px 0 0 0;max-height:220px;overflow:auto"></ul>
      `;
      div.style.height = '300px';
      div.style.padding = '6px 10px';

      const ul = bodyEl.querySelector('#searchList');
      const searchInput = bodyEl.querySelector('#pileSearch');

      const allPiles = markers.map(m => {
        const code = extractPileCode(m.name);
        const s = code ? stockIndexGlobal[code] : null;
        const typeLabel = (markerConfig[m.type] && markerConfig[m.type].displayName) ? markerConfig[m.type].displayName : (m.type || '');
        return {
          code,
          name: m.name,
          type: typeLabel,       // display name
          rawType: m.type,       // <-- add this raw type key
          material: (s && s.material) ? s.material : '',
          marker: m._leaflet
        };
      });

      function renderList(filterText = '') {
        ul.innerHTML = '';
        const term = filterText.trim().toLowerCase();
        const items = allPiles.filter(p => {
          if (!term) return true;
          return (
            (p.code && p.code.toLowerCase().includes(term)) ||
            (p.name && p.name.toLowerCase().includes(term)) ||
            (p.material && p.material.toLowerCase().includes(term)) ||
            (p.type && p.type.toLowerCase().includes(term))
          );
        }).sort((a,b) => {
          const ac = a.code || ''; const bc = b.code || '';
          if (ac !== bc) return ac.localeCompare(bc);
          return a.name.localeCompare(b.name);
        });

        items.forEach(p => {
          const li = document.createElement('li');
          li.style.padding = '6px 0';
          li.style.borderBottom = '1px dashed #eee';
          li.style.cursor = 'default';
          const sub = p.material ? `${p.type} — ${p.material}` : p.type;
          li.innerHTML = `
            <div style="display:flex;justify-content:space-between;align-items:center">
              <div>
                <div style="font-weight:600">${p.code || '—'} — ${sub}</div>
                <div style="color:#555">${p.name}</div>
              </div>
              <button style="margin-left:8px;padding:2px 6px;border:1px solid #ccc;border-radius:3px;cursor:pointer">Ping</button>
            </div>
          `; 
          li.addEventListener('click', () => {
            const target =
              (p.rawType === 'Coils' && window.burningArea) ? window.burningArea :
              ((p.rawType === 'Breaking' || p.rawType === 'Unbreakable') && window.breakingArea) ? window.breakingArea :
              p.marker;
            if (target) pingMarker(target);
          });      
          li.querySelector('button').addEventListener('click', (e) => {
            e.stopPropagation();
            const target =
              (p.rawType === 'Coils' && window.burningArea) ? window.burningArea :
              ((p.rawType === 'Breaking' || p.rawType === 'Unbreakable') && window.breakingArea) ? window.breakingArea :
              p.marker;
            if (target) pingMarker(target);
          });
          ul.appendChild(li);
        });
      }
      renderList();
      searchInput.addEventListener('input', (e) => renderList(e.target.value));
    }

    renderHeader();
    renderBody();
    headerEl.addEventListener('click', (e) => {
      const btn = headerEl.querySelector('#srchToggleBtn');
      if (btn && e.target === btn) { /* keep default */ }
     const wasCollapsed = collapsed;
     collapsed = !collapsed;
     renderHeader();
     renderBody();
     // If we just expanded, collapse the other panel
     if (!collapsed) collapseOthers('searchPanel');
    });

    div.appendChild(headerEl);
    div.appendChild(bodyEl);
     registerPanel('searchPanel', { isExpanded, expand, collapse });
    return div;
  };
    searchCtrl.addTo(map);
    pastDueCtrl.addTo(map);
    if (unknownTypes.size) console.warn('Unknown types:', Array.from(unknownTypes));
}).catch(err => console.error('Data load failed:', err));

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
``
