// Historical Dashboard — month drill-down overlay.
// Daily bar chart with a metric filter and filter-specific viewers:
//  - Heats → expanded heat viewer (heat → grade → charge buckets → materials)
//  - Received (Total) → truck viewer (expandable per-truck detail) + railcar viewer
// Fed by the same month-cached fetchers the map popups use.
// Requires: utils.js, DashboardCompare.js (dashBarChart, dashNum), and the
// five station files. Loaded before Dashboard.js.

const DASH_DAILY_METRICS = [
  { key: 'consumed', label: 'Consumed',       unit: 'tons',   dec: 0 },
  { key: 'heats',    label: 'Heats',          unit: 'heats',  dec: 0 },
  { key: 'received', label: 'Received',       unit: 'trucks', dec: 0 },
  { key: 'broken',   label: 'Broken',         unit: 'tons',   dec: 0 },
  { key: 'billable', label: 'Cut — Billable', unit: 'tons',   dec: 1 },
];

// Overlay state: null when closed.
let _dashOvl = null;

function dashDayCell(year, month, day) {
  const dow = new Date(year, month, day).getDay();
  const letters = ['S', 'M', 'T', 'W', 'T', 'F', 'S'];
  const color = dow === 0 ? '#f87171' : dow === 6 ? '#fb923c' : dow === 3 ? '#4ade80' : '#60a5fa';
  return `<span class="dash-wk" style="color:${color}">${letters[dow]}</span>${month + 1}/${day}`;
}

function dashIsoDay(iso) {
  const m = String(iso || '').match(/^\d{4}-\d{1,2}-(\d{1,2})$/);
  return m ? +m[1] : null;
}

// ─── Data model ────────────────────────────────────────────────────────
// Merges the five month payloads into one per-day array plus the raw row
// lists each activity table needs.
function buildDashMonthModel(year, month, src) {
  const daysInMonth = new Date(year, month + 1, 0).getDate();
  const days = Array.from({ length: daysInMonth }, (_, i) => ({
    day: i + 1,
    consumedTons: 0, heats: 0, trucks: 0, truckLbs: 0, brokenTons: 0, billableTons: 0,
    hasData: false,
  }));
  const at = d => (d >= 1 && d <= daysInMonth) ? days[d - 1] : null;

  (src.bucket?.rows || []).forEach(r => {
    const t = at(r.day || dashIsoDay(r.isoDate));
    if (!t) return;
    t.consumedTons += r.tons || 0;
    t.heats += r.heatsCompleted || 0;
    t.hasData = true;
  });
  (src.recv?.rows || []).forEach(r => {
    const t = at(dashIsoDay(r.isoDate));
    if (!t) return;
    t.trucks += r.trucks || 0;
    t.truckLbs += r.weight || 0;
    t.hasData = true;
  });
  (src.breaking?.rows || []).forEach(r => {
    const t = (r.date instanceof Date) ? at(r.date.getDate()) : null;
    if (!t) return;
    t.brokenTons += r.netTons || 0;
    t.hasData = true;
  });
  (src.burning?.rows || []).forEach(r => {
    const t = (r.date instanceof Date) ? at(r.date.getDate()) : null;
    if (!t) return;
    t.billableTons += r.billableTons || 0;
    t.hasData = true;
  });

  // Current month is partial: only chart up to the last day with data.
  const { year: curY, month: curM } = getCurrentInventoryPeriod();
  let lastDay = daysInMonth;
  if (year === curY && month === curM) {
    lastDay = 0;
    days.forEach(d => { if (d.hasData) lastDay = d.day; });
    if (lastDay === 0) lastDay = Math.min(new Date().getDate(), daysInMonth);
  }

  // Railcars have no release date in the source sheet — monthly list only.
  const released = src.rail
    ? (src.rail.released?.length ? src.rail.released : (src.rail.railcars || []))
    : [];
  const cars = released.map(c => {
    const g = parseWeight(c.ourGross), t = parseWeight(c.ourTare);
    let net = (Number.isFinite(g) && Number.isFinite(t)) ? g - t : parseWeight(c.shipperNet);
    if (!Number.isFinite(net)) net = null;
    return { car: c.railcarNum, supplier: c.supplier, material: c.material, lot: c.lot, netLbs: net, status: c.status || '—' };
  });
  const railLbs = cars.reduce((s, c) => s + (c.netLbs || 0), 0);

  const rowDay = r => r.day || dashIsoDay(r.isoDate) || 0;
  const sortAsc = rows => [...rows].sort((a, b) => a.date - b.date);
  return {
    days: days.slice(0, lastDay),
    bucketRows: [...(src.bucket?.rows || [])].sort((a, b) => rowDay(a) - rowDay(b)),
    recvRows: src.recv?.rows || [],
    breakingRows: sortAsc(src.breaking?.rows || []),
    burningRows: sortAsc(src.burning?.rows || []),
    cars, railLbs,
    totals: {
      bucket: src.bucket?.totals || null,
      recv: src.recv?.totals || null,
      breaking: src.breaking?.totals || null,
      burning: src.burning?.totals || null,
    },
  };
}

// ─── Open / close ──────────────────────────────────────────────────────
// metric selects the initial view (e.g. 'heats' from the Heats MoM chart,
// 'received' from any Received-family MoM chart).
async function openDashMonth(year, month, metric = 'consumed') {
  if (!DASH_DAILY_METRICS.some(m => m.key === metric)) metric = 'consumed';
  _dashOvl = { year, month, metric, recvSub: 'both', openDay: null, openHeat: null, openTruckDay: null, model: null };
  const ovl = document.getElementById('dashOvl');
  const title = document.getElementById('dashOvlTitle');
  const body = document.getElementById('dashOvlBody');
  if (!ovl || !title || !body) return;

  const { year: curY, month: curM } = getCurrentInventoryPeriod();
  title.textContent = formatMonthYear(year, month) + (year === curY && month === curM ? ' · month to date' : '');
  body.innerHTML = '<div style="color:#64748b;font-size:13px;padding:24px 4px">Loading month data…</div>';
  ovl.style.display = 'block';

  const override = { year, month };
  const [bucket, recv, burning, breaking, rail] = await Promise.all([
    fetchBucketLoadingConsumption(false, override).catch(() => null),
    fetchReceivingSummary(false, override).catch(() => null),
    fetchBurningTotals(false, override).catch(() => null),
    fetchBreakingTotals(false, override).catch(() => null),
    fetchRailcarSummary(false, override).catch(() => null),
  ]);

  // Bail if the overlay was closed or retargeted while loading.
  if (!_dashOvl || _dashOvl.year !== year || _dashOvl.month !== month) return;

  if (!bucket && !recv && !burning && !breaking && !rail) {
    body.innerHTML = '<div style="color:#64748b;font-size:13px;padding:24px 4px">No data found for this month.</div>';
    return;
  }

  _dashOvl.model = buildDashMonthModel(year, month, { bucket, recv, burning, breaking, rail });
  renderDashOvl();
}

function closeDashMonth() {
  const ovl = document.getElementById('dashOvl');
  if (ovl) ovl.style.display = 'none';
  _dashOvl = null;
}

// ─── Activity tables ───────────────────────────────────────────────────
function dashConsumptionTable() {
  const { year, month, model } = _dashOvl;
  const t = model.totals.bucket || {};
  const totalBuckets = model.bucketRows.reduce((s, r) => s + (r.bucketsLoaded || 0), 0);
  // The heat viewer only opens from the month-specific Heats view.
  const showHeatViewer = _dashOvl.metric === 'heats';

  const dayRows = model.bucketRows.map(r => {
    const day = r.day || dashIsoDay(r.isoDate);
    const heats = Array.isArray(r.heatBreakdown) ? r.heatBreakdown : [];
    const open = showHeatViewer && _dashOvl.openDay === r.isoDate;
    let detail = '';
    if (open && heats.length) {
      const heatRows = heats.map((h, hi) => {
        const mats = Array.isArray(h.materials) ? h.materials : [];
        const totLbs = mats.reduce((s, x) => s + (Number(x.pounds) || 0), 0);
        const totTons = mats.reduce((s, x) => s + (Number(x.tons) || 0), 0);
        const hOpen = _dashOvl.openHeat === hi;
        const matRows = hOpen
          ? `<tr class="dash-mat-row" style="background:#233150"><td style="color:#64748b;font-size:11px">Bucket</td><td style="color:#64748b;font-size:11px">Pile</td><td style="color:#64748b;font-size:11px">Material</td><td style="color:#64748b;font-size:11px">Lot #</td><td style="color:#64748b;font-size:11px">Pounds</td><td style="color:#64748b;font-size:11px">Tons</td></tr>`
            + (mats.length
              ? mats.map(x => {
                  const bktColor = bucketColorForSeq(x.bucketSeq);
                  const material = x.material || getMaterialForLotOrPile(x.lot, x.pile);
                  return `<tr class="dash-mat-row"${bktColor ? ` style="background:${bktColor}1c"` : ''}><td>${bucketBadgeHtml(x)}</td><td>${esc(x.pile || '—')}</td><td>${esc(material || '—')}</td><td>${esc(x.lot || '—')}</td><td>${fmtInt(x.pounds)}</td><td>${fmtTons2(x.tons, 2)}</td></tr>`;
                }).join('')
              : '<tr class="dash-mat-row"><td colspan="6" style="color:#64748b;font-style:italic">No material records</td></tr>')
          : '';
        return `<tr class="dash-heat-hdr${hOpen ? ' open' : ''}" data-heat="${hi}">
          <td><span class="chev">▶</span><b>${esc(h.heatNumber)}</b></td><td>${esc(h.grade || '—')}</td><td>${h.bucketCount ? fmtInt(h.bucketCount) : '—'}</td><td></td>
          <td>${fmtInt(totLbs)}</td><td>${fmtTons2(totTons, 2)}</td></tr>${matRows}`;
      }).join('');
      detail = `<tr class="dash-detail"><td colspan="5">
        <div style="font-size:12px;color:#94a3b8;margin-bottom:6px"><b style="color:#e2e8f0">${heats.length} heat${heats.length !== 1 ? 's' : ''} completed on ${month + 1}/${day}</b> — click a heat to see its charge buckets and materials</div>
        <table class="dash-heat-tbl"><tr><th>Heat #</th><th>Grade</th><th>Buckets</th><th></th><th>Total Lbs</th><th>Total Tons</th></tr>${heatRows}</table>
      </td></tr>`;
    }
    return `<tr>
      <td>${day ? dashDayCell(year, month, day) : esc(r.dateLabel)}</td>
      <td>${fmtInt(r.pounds)}</td>
      <td>${fmtTons2(r.tons, 2)}</td>
      <td>${fmtInt(r.bucketsLoaded)}</td>
      <td>${showHeatViewer && heats.length > 0
        ? `<button type="button" class="dash-heat-btn${open ? ' active' : ''}" data-date="${esc(r.isoDate)}">${fmtInt(r.heatsCompleted)}</button>`
        : fmtInt(r.heatsCompleted)}</td>
    </tr>${detail}`;
  }).join('');

  const hint = showHeatViewer
    ? ' <span style="font-weight:400;text-transform:none;letter-spacing:0">(click a heat count to open the heat viewer)</span>'
    : '';
  return `<div class="dash-sec-title">Bucket Loading — Daily Activity${hint}</div>
    <div class="dash-tbl-wrap"><table>
      <tr><th>Date</th><th>Pounds</th><th>Tons</th><th>Buckets</th><th>Heats</th></tr>
      ${dayRows || '<tr><td colspan="5" style="color:#64748b">No consumption rows for this month.</td></tr>'}
      <tr class="total-row"><td>Total</td><td>${fmtInt(t.totalPounds)}</td><td>${fmtTons2(t.totalTons, 2)}</td><td>${fmtInt(totalBuckets)}</td><td>${fmtInt(t.totalHeats)}</td></tr>
    </table></div>`;
}

function dashTruckTable() {
  const { year, month, model } = _dashOvl;
  const t = model.totals.recv || {};
  const pastDueCodes = new Set((window.pastDue || []).map(p => normalizePileCode(p.code)));

  const rows = model.recvRows.map(r => {
    const day = dashIsoDay(r.isoDate);
    const trucks = Array.isArray(r.truckDetails) ? r.truckDetails : [];
    const open = _dashOvl.openTruckDay === r.isoDate;
    let detail = '';
    if (open && trucks.length) {
      const truckRows = trucks.map(item => {
        const material = item.material || getMaterialForLotOrPile(item.lot, item.pile);
        const flagged = item.isStoppedPile || pastDueCodes.has(normalizePileCode(item.pile || ''));
        const flagStyle = flagged ? ' style="color:#f87171;font-weight:700"' : '';
        const net = parseWeight(item.net);
        const infoBits = [];
        if (item.gross || item.tare) infoBits.push(`Gross: ${esc(item.gross || '—')} · Tare: ${esc(item.tare || '—')}`);
        if (item.remarks) infoBits.push(esc(item.remarks));
        const infoRow = infoBits.length
          ? `<tr class="dash-mat-row"><td colspan="5" style="text-align:left;color:#94a3b8;font-size:11px;white-space:normal">${infoBits.join(' &nbsp;—&nbsp; ')}</td></tr>`
          : '';
        return `<tr class="dash-mat-row">
          <td>${esc(item.truckNumber || '—')}</td>
          <td>${esc(item.ticketId || '—')}</td>
          <td${flagStyle}>${esc(item.pile || '—')}</td>
          <td class="l"${flagStyle}>${esc(material || '—')}</td>
          <td>${Number.isFinite(net) ? fmtInt(net) : '—'}</td>
        </tr>${infoRow}`;
      }).join('');
      detail = `<tr class="dash-detail"><td colspan="4">
        <div style="font-size:12px;color:#94a3b8;margin-bottom:6px"><b style="color:#e2e8f0">${fmtInt(r.trucks)} trucks received on ${month + 1}/${day}</b></div>
        <table class="dash-heat-tbl"><tr><th>Truck #</th><th>Ticket ID</th><th>Pile</th><th class="l">Material</th><th>Net Lbs</th></tr>${truckRows}</table>
      </td></tr>`;
    }
    return `<tr>
      <td>${day ? dashDayCell(year, month, day) : esc(r.dateLabel)}</td>
      <td>${trucks.length > 0
        ? `<button type="button" class="dash-truck-btn${open ? ' active' : ''}" data-date="${esc(r.isoDate)}">${fmtInt(r.trucks)}</button>`
        : fmtInt(r.trucks)}</td>
      <td>${fmtInt(r.weight)}</td>
      <td>${fmtTons2((r.weight || 0) / 2000, 2)}</td>
    </tr>${detail}`;
  }).join('');
  return `<div class="dash-sec-title">Trucks Received — Daily Activity <span style="font-weight:400;text-transform:none;letter-spacing:0">(click a truck count to open the truck viewer)</span></div>
    <div class="dash-tbl-wrap"><table>
      <tr><th>Date</th><th>Trucks</th><th>Pounds</th><th>Tons</th></tr>
      ${rows || '<tr><td colspan="4" style="color:#64748b">No truck receiving rows for this month.</td></tr>'}
      <tr class="total-row"><td>Total</td><td>${fmtInt(t.totalTrucks)}</td><td>${fmtInt(t.totalWeight)}</td><td>${fmtTons2(t.totalTons, 2)}</td></tr>
    </table></div>`;
}

function dashRailTable() {
  const { model } = _dashOvl;
  const rows = model.cars.map(c => {
    const statusLower = String(c.status).toLowerCase();
    const statusColor = statusLower.includes('released') ? '#22c55e'
      : statusLower.includes('reject') ? '#f87171' : '#94a3b8';
    const material = c.material || getMaterialForLotOrPile(c.lot, '');
    return `<tr>
      <td>${esc(c.car)}</td>
      <td class="l">${esc(c.supplier || '—')}</td>
      <td class="l">${esc(material || '—')}</td>
      <td>${esc(c.lot || '—')}</td>
      <td>${c.netLbs !== null ? fmtInt(c.netLbs) : '—'}</td>
      <td>${c.netLbs !== null ? fmtTons2(c.netLbs / 2000, 2) : '—'}</td>
      <td class="l" style="color:${statusColor}">${esc(c.status)}</td>
    </tr>`;
  }).join('');
  return `<div class="dash-sec-title">Railcars Received — ${fmtInt(model.cars.length)} cars this month</div>
    <div class="dash-note">Railcar releases aren't dated in the source sheet (Car Status only), so railcars are compared month-to-month and listed here rather than charted by day.</div>
    <div class="dash-tbl-wrap" style="max-height:340px;overflow-y:auto"><table>
      <tr><th>Car</th><th class="l">Supplier</th><th class="l">Material</th><th>Lot #</th><th>Net Lbs</th><th>Net Tons</th><th class="l">Status</th></tr>
      ${rows || '<tr><td colspan="7" style="color:#64748b">No railcars recorded for this month.</td></tr>'}
      <tr class="total-row"><td>Total</td><td class="l"></td><td class="l"></td><td></td><td>${fmtInt(model.railLbs)}</td><td>${fmtTons2(model.railLbs / 2000, 2)}</td><td class="l"></td></tr>
    </table></div>`;
}

function dashBreakingTable() {
  const { year, month, model } = _dashOvl;
  const t = model.totals.breaking || {};
  const rows = model.breakingRows.map(r => {
    const day = (r.date instanceof Date) ? r.date.getDate() : null;
    return `<tr>
      <td>${day ? dashDayCell(year, month, day) : esc(r.dateLabel)}</td>
      <td class="l">${esc(r.material || '—')}</td>
      <td>${fmtInt(r.netLbs)}</td>
      <td>${fmtTons2(r.netTons, 2)}</td>
    </tr>`;
  }).join('');
  return `<div class="dash-sec-title">Breaking — Daily Activity</div>
    <div class="dash-tbl-wrap"><table>
      <tr><th>Date</th><th class="l">Material</th><th>Net Lbs</th><th>Net Tons</th></tr>
      ${rows || '<tr><td colspan="4" style="color:#64748b">No breaking rows for this month.</td></tr>'}
      <tr class="total-row"><td>Total</td><td class="l"></td><td>${fmtInt(t.totalLbs)}</td><td>${fmtTons2(t.totalTons, 2)}</td></tr>
    </table></div>`;
}

function dashBurningTable() {
  const { year, month, model } = _dashOvl;
  const t = model.totals.burning || {};
  const rows = model.burningRows.map(r => {
    const day = (r.date instanceof Date) ? r.date.getDate() : null;
    return `<tr>
      <td>${day ? dashDayCell(year, month, day) : esc(r.dateLabel)}</td>
      <td><b>${fmtTons2(r.billableTons, 2)}</b></td>
      <td>${fmtTons2(r.netTons, 2)}</td>
      <td>${fmtInt(r.cuts)}</td>
    </tr>`;
  }).join('');
  return `<div class="dash-sec-title">Torch Cutting — Daily Activity <span style="font-weight:400;text-transform:none;letter-spacing:0">(billable tons first)</span></div>
    <div class="dash-tbl-wrap"><table>
      <tr><th>Date</th><th>Billable Tons</th><th>Total Cut Tons</th><th>Cuts</th></tr>
      ${rows || '<tr><td colspan="4" style="color:#64748b">No cutting rows for this month.</td></tr>'}
      <tr class="total-row"><td>Total</td><td>${fmtTons2(t.billableTons, 2)}</td><td>${fmtTons2(t.netTons, 2)}</td><td>${fmtInt(t.cuts)}</td></tr>
    </table></div>`;
}

// ─── Overlay rendering ─────────────────────────────────────────────────
function renderDashOvl() {
  if (!_dashOvl || !_dashOvl.model) return;
  const { year, month, model } = _dashOvl;
  const body = document.getElementById('dashOvlBody');
  if (!body) return;

  const dm = DASH_DAILY_METRICS.find(x => x.key === _dashOvl.metric) || DASH_DAILY_METRICS[0];
  const isRecv = dm.key === 'received';
  const getVal = d => (
    dm.key === 'consumed' ? d.consumedTons :
    dm.key === 'heats' ? d.heats :
    dm.key === 'received' ? d.trucks :
    dm.key === 'broken' ? d.brokenTons : d.billableTons
  );

  const chips = DASH_DAILY_METRICS.map(x =>
    `<button type="button" class="dash-chip${x.key === dm.key ? ' on1' : ''}" data-dkey="${x.key}"><span class="dot"></span>${x.label}</button>`).join('');

  const subChips = isRecv
    ? `<div class="dash-sub-chips"><span class="sub-lbl">Show</span>`
      + [['both', 'Trucks + Railcars'], ['trucks', 'Trucks only'], ['rail', 'Railcars only']].map(([v, lbl]) =>
        `<button type="button" class="dash-chip sm${_dashOvl.recvSub === v ? ' on1' : ''}" data-rsub="${v}"><span class="dot"></span>${lbl}</button>`).join('')
      + '</div>'
    : '';

  // Stats row
  let stats;
  if (isRecv) {
    const tR = model.totals.recv || {};
    const railTons = model.railLbs / 2000;
    const combined = (tR.totalTons || 0) + railTons;
    stats = `
      <div class="dash-stat"><div class="s-lbl">Trucks Received</div><div class="s-val">${fmtInt(tR.totalTrucks)} <span class="unit">trucks · ${dashNum(tR.totalTons, 0)} tons</span></div></div>
      <div class="dash-stat"><div class="s-lbl">Railcars Received</div><div class="s-val">${fmtInt(model.cars.length)} <span class="unit">cars · ${dashNum(railTons, 0)} tons</span></div></div>
      <div class="dash-stat"><div class="s-lbl">Received Total</div><div class="s-val">${dashNum(combined, 0)} <span class="unit">tons</span></div></div>`;
  } else {
    const vals = model.days.map(getVal);
    const nz = vals.filter(v => v > 0);
    const total = vals.reduce((a, b) => a + b, 0);
    const peak = model.days.reduce((a, b) => (getVal(b) > getVal(a) ? b : a), model.days[0]);
    stats = `
      <div class="dash-stat"><div class="s-lbl">Total ${esc(dm.label)}</div><div class="s-val">${dashNum(total, dm.dec)} <span class="unit">${esc(dm.unit)}</span></div></div>
      <div class="dash-stat"><div class="s-lbl">Avg / operating day</div><div class="s-val">${dashNum(total / (nz.length || 1), dm.dec)} <span class="unit">${esc(dm.unit)}</span></div></div>
      <div class="dash-stat"><div class="s-lbl">Peak day</div><div class="s-val">${peak ? `${month + 1}/${peak.day}` : '—'} <span class="unit">${peak ? `${dashNum(getVal(peak), dm.dec)} ${esc(dm.unit)}` : ''}</span></div></div>`;
  }

  // Daily chart — value-only tooltips; railcars have no daily dates to chart.
  let chart = '';
  if ((!isRecv || _dashOvl.recvSub !== 'rail') && model.days.length) {
    const letters = ['S', 'M', 'T', 'W', 'T', 'F', 'S'];
    const chartVals = model.days.map(d => ({
      label: `${letters[new Date(year, month, d.day).getDay()]} ${month + 1}/${d.day}`,
      axis: String(d.day),
      value: getVal(d),
      hl: false,
    }));
    chart = dashBarChart({
      values: chartVals, color: '#3987e5', height: 170,
      dec: dm.dec, unit: isRecv ? 'trucks' : dm.unit,
      labelEvery: model.days.length > 12 ? 2 : 1,
    });
    if (isRecv) chart = `<div class="dash-sec-title" style="margin-top:4px">Trucks received per day</div>` + chart;
  }

  // Filter-specific tables
  let tables = '';
  if (dm.key === 'consumed' || dm.key === 'heats') tables = dashConsumptionTable();
  else if (isRecv) {
    if (_dashOvl.recvSub === 'both') tables = dashTruckTable() + dashRailTable();
    else if (_dashOvl.recvSub === 'trucks') tables = dashTruckTable();
    else tables = dashRailTable();
  }
  else if (dm.key === 'broken') tables = dashBreakingTable();
  else if (dm.key === 'billable') tables = dashBurningTable();

  body.innerHTML = `
    <div class="dash-ovl-chips">${chips}</div>
    ${subChips}
    <div class="dash-ovl-stats">${stats}</div>
    ${chart}
    ${tables}`;
}

// ─── Event wiring (delegation) ─────────────────────────────────────────
document.getElementById('dashOvlBody')?.addEventListener('click', e => {
  const chip = e.target.closest('.dash-chip[data-dkey]');
  if (chip) {
    _dashOvl.metric = chip.getAttribute('data-dkey');
    _dashOvl.openDay = null;
    _dashOvl.openHeat = null;
    _dashOvl.openTruckDay = null;
    renderDashOvl();
    return;
  }
  const sub = e.target.closest('.dash-chip[data-rsub]');
  if (sub) {
    _dashOvl.recvSub = sub.getAttribute('data-rsub');
    renderDashOvl();
    return;
  }
  const tb = e.target.closest('.dash-truck-btn');
  if (tb) {
    const key = tb.getAttribute('data-date');
    _dashOvl.openTruckDay = _dashOvl.openTruckDay === key ? null : key;
    renderDashOvl();
    return;
  }
  const hb = e.target.closest('.dash-heat-btn');
  if (hb) {
    const key = hb.getAttribute('data-date');
    _dashOvl.openHeat = null;
    _dashOvl.openDay = _dashOvl.openDay === key ? null : key;
    renderDashOvl();
    return;
  }
  const hh = e.target.closest('.dash-heat-hdr');
  if (hh) {
    const hi = +hh.getAttribute('data-heat');
    _dashOvl.openHeat = _dashOvl.openHeat === hi ? null : hi;
    renderDashOvl();
  }
});

document.getElementById('dashOvlClose')?.addEventListener('click', closeDashMonth);
document.getElementById('dashOvl')?.addEventListener('click', e => {
  if (e.target.id === 'dashOvl') closeDashMonth();
});
document.addEventListener('keydown', e => {
  if (e.key === 'Escape' && _dashOvl) closeDashMonth();
});
