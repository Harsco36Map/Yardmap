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
const consumptionCsvUrl = 'AverageConsumption.csv'; // <-- set per month

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
    div.innerHTML = '<div>Inventory data current as of —</div><div style="margin-top:4px;color:#555">Total yard inventory: —</div><div id="materialFlowLines"><div id="materialReceivedBannerLine" style="margin-top:4px;padding:1px 4px;border-radius:3px;background:#c8e6c9;color:#1b5e20">Material Received: —</div><div id="materialConsumedBannerLine" style="margin-top:4px;padding:1px 4px;border-radius:3px;background:#ffcdd2;color:#b71c1c">Material Consumed: —</div></div>';
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
        <div style="margin-top:4px;color:#555">Total yard inventory: <b id="invWeightSpan" style="cursor:help">${inventoryText}</b></div>
        <div id="materialFlowLines">
          <div id="materialReceivedBannerLine" style="margin-top:4px;padding:1px 4px;border-radius:3px;background:#c8e6c9;color:#1b5e20">Material Received: —</div>
          <div id="materialConsumedBannerLine" style="margin-top:4px;padding:1px 4px;border-radius:3px;background:#ffcdd2;color:#b71c1c">Material Consumed: —</div>
        </div>
      `;
      updateMaterialReceivedBanner();
    }
  }).catch(err => console.warn('stockData.json meta fetch failed:', err));
})();


function updateMaterialReceivedBanner() {
  const receivedEl = document.getElementById('materialReceivedBannerLine');
  const consumedEl = document.getElementById('materialConsumedBannerLine');
  const flowLines = document.getElementById('materialFlowLines');
  const invSpan = document.getElementById('invWeightSpan');

  const hasReceived = currentMonthTruckWeight !== null || currentMonthRailWeight !== null;
  const hasConsumed = currentMonthBucketPounds !== null;

  const truck = typeof currentMonthTruckWeight === 'number' ? currentMonthTruckWeight : 0;
  const rail = typeof currentMonthRailWeight === 'number' ? currentMonthRailWeight : 0;
  const received = truck + rail;
  const consumed = typeof currentMonthBucketPounds === 'number' ? currentMonthBucketPounds : 0;

  if (receivedEl && hasReceived) {
    receivedEl.innerHTML = `Material Received: <b>${fmtInt(received)} lbs</b>`;
  }
  if (consumedEl && hasConsumed) {
    consumedEl.innerHTML = `Material Consumed: <b>${fmtInt(consumed)} lbs</b>`;
  }

  if (flowLines && receivedEl && consumedEl && hasReceived && hasConsumed) {
    if (consumed > received) {
      flowLines.insertBefore(consumedEl, receivedEl);
    } else {
      flowLines.insertBefore(receivedEl, consumedEl);
    }
    if (invSpan) {
      const net = received - consumed;
      const netText = `${net >= 0 ? '+' : '-'} ${fmtInt(Math.abs(net))} lbs`;
      invSpan.removeAttribute('title');

      let tip = document.getElementById('invNetTooltip');
      if (!tip) {
        tip = document.createElement('div');
        tip.id = 'invNetTooltip';
        tip.style.cssText = 'display:none;position:fixed;background:#333;color:#fff;padding:3px 8px;border-radius:3px;font:12px system-ui,sans-serif;pointer-events:none;z-index:9999;white-space:nowrap';
        document.body.appendChild(tip);
      }
      tip.textContent = netText;

      invSpan.onmouseenter = () => {
        const bannerRect = document.getElementById('invBanner').getBoundingClientRect();
        tip.style.left = bannerRect.left + 'px';
        tip.style.top  = (bannerRect.bottom + 4) + 'px';
        tip.style.display = 'block';
      };
      invSpan.onmousemove  = null;
      invSpan.onmouseleave = () => { tip.style.display = 'none'; };
    }
  }
}

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
let currentMonthTruckWeight = null;
let currentMonthRailWeight = null;
let currentMonthBucketPounds = null;
let historyWorkbookCache = { at: 0, workbook: null };
let latestInventoryPeriod = { year: null, month: null };

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

  // Auto-fetch banner data after all initialization is complete so the XLSX
  // parse of Production.xlsx does not block stockIndexGlobal from being built.
  (async () => {
    try { await fetchReceivingSummary(); } catch (e) { console.warn('Auto-fetch receiving failed:', e); }
    try { await fetchRailcarSummary(); } catch (e) { console.warn('Auto-fetch railcar failed:', e); }
    try { await fetchBucketLoadingConsumption(); } catch (e) { console.warn('Auto-fetch bucket loading failed:', e); }
  })();
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

