// Aggregates all five data sources for a single month into one summary object.
// Uses the existing fetch functions with periodOverride so no global state is mutated.
async function fetchMonthSummary(year, month) {
  const override = { year, month };
  const [recv, bucket, burning, breaking, rail] = await Promise.all([
    fetchReceivingSummary(false, override).catch(() => null),
    fetchBucketLoadingConsumption(false, override).catch(() => null),
    fetchBurningTotals(false, override).catch(() => null),
    fetchBreakingTotals(false, override).catch(() => null),
    fetchRailcarSummary(false, override).catch(() => null),
  ]);

  const released = rail?.released ?? [];
  const railWeightLbs = released.reduce((sum, car) => {
    const g = parseWeight(car.ourGross);
    const t = parseWeight(car.ourTare);
    return (Number.isFinite(g) && Number.isFinite(t)) ? sum + (g - t) : sum;
  }, 0);

  const truckLbs = recv?.totals?.totalWeight ?? null;
  const railLbs  = railWeightLbs > 0 ? railWeightLbs : (released.length > 0 ? 0 : null);

  return {
    truckTons:    recv?.totals?.totalTons    ?? null,
    truckLbs,
    truckCount:   recv?.totals?.totalTrucks  ?? null,
    railTons:     railLbs !== null ? railLbs / 2000 : null,
    railLbs,
    railcarCount: released.length > 0 || rail !== null ? released.length : null,
    consumedTons: bucket?.totals?.totalTons   ?? null,
    consumedLbs:  bucket?.totals?.totalPounds ?? null,
    totalHeats:   bucket?.totals?.totalHeats  ?? null,
    avgDailyTons: bucket?.totals?.avgTons     ?? null,
    brokenTons:    breaking?.totals?.totalTons  ?? null,
    brokenLbs:     breaking?.totals?.totalLbs   ?? null,
    cutTons:       burning?.totals?.netTons     ?? null,
    cutLbs:        burning?.totals?.netLbs      ?? null,
    billableTons:  burning?.totals?.billableTons ?? null,
    billableLbs:   burning?.totals?.billableTons != null ? burning.totals.billableTons * 2000 : null,
  };
}

// Returns true if any field in a summary object has actual data.
function dashSummaryHasData(s) {
  return s && Object.values(s).some(v => v !== null && v !== 0);
}

// Formats a number as tons with one decimal place, or '—' if null.
function dashFmt(val, decimals = 1) {
  if (val === null || val === undefined || !Number.isFinite(val)) return '—';
  return val.toLocaleString('en-US', { minimumFractionDigits: decimals, maximumFractionDigits: decimals });
}

// Renders the HTML for one month card. data is null for months with no history.
function renderDashCard(year, month, data) {
  const MONTHS = ['January','February','March','April','May','June','July','August','September','October','November','December'];
  const name = MONTHS[month] || 'Month';
  const { year: curY, month: curM } = getCurrentInventoryPeriod();
  const isCurrent = year === curY && month === curM;
  const cardClass = 'dash-card' + (isCurrent ? ' dash-card--current' : '');
  const titleClass = 'dash-card-title' + (isCurrent ? ' current' : '');

  if (!data || !dashSummaryHasData(data)) {
    return `<div class="${cardClass} dash-card--empty">
      <div>
        <div class="${titleClass}">${name}</div>
        <div style="margin-top:12px">No data</div>
      </div>
    </div>`;
  }

  const truckRailTons = (data.truckTons ?? 0) + (data.railTons ?? 0);
  const netChange = (truckRailTons > 0 || data.consumedTons !== null)
    ? truckRailTons - (data.consumedTons ?? 0)
    : null;
  const netChangeLbs = (data.truckLbs !== null || data.railLbs !== null || data.consumedLbs !== null)
    ? ((data.truckLbs ?? 0) + (data.railLbs ?? 0)) - (data.consumedLbs ?? 0)
    : null;
  const netClass = netChange === null ? '' : (netChange >= 0 ? ' positive' : ' negative');

  const row = (label, value, unit = '') =>
    `<div class="dash-metric">
      <span class="dash-metric-label">${label}</span>
      <span class="dash-metric-value">${value}${unit ? '<span style="color:#475569;font-weight:400"> ' + unit + '</span>' : ''}</span>
    </div>`;

  const fmtLbs = rawLbs => (rawLbs !== null && Number.isFinite(rawLbs))
    ? rawLbs.toLocaleString('en-US', { maximumFractionDigits: 0 }) + ' lbs'
    : null;

  const truckLbs    = fmtLbs(data.truckLbs);
  const railLbs     = fmtLbs(data.railLbs);
  const consumedLbs = fmtLbs(data.consumedLbs);

  const combinedRow = (label, tons, lbs, count, countUnit) => {
    const tonsStr = dashFmt(tons);
    const weight = lbs ? `${tonsStr} tons / ${lbs}` : `${tonsStr} tons`;
    const countStr = count !== null ? `${count} ${countUnit}` : '';
    return `<div class="dash-metric" style="align-items:flex-start">
      <span class="dash-metric-label" style="padding-top:1px">${label}</span>
      <span class="dash-metric-value" style="text-align:right">
        ${weight}
        ${countStr ? `<br><span style="color:#64748b;font-size:10px;font-weight:400">${countStr}</span>` : ''}
      </span>
    </div>`;
  };

  return `<div class="${cardClass}">
    <div class="${titleClass}">${name}${isCurrent ? ' <span style="font-size:9px;vertical-align:middle;opacity:0.8">●</span>' : ''}</div>
    ${combinedRow('Consumed', data.consumedTons, consumedLbs, data.totalHeats !== null && data.totalHeats > 0 ? data.totalHeats : null, 'heats')}
    ${row('Avg Daily Use', dashFmt(data.avgDailyTons), 'tons/day')}
    <hr class="dash-metric-divider">
    ${combinedRow('Broken', data.brokenTons, fmtLbs(data.brokenLbs), null, '')}
    <div class="dash-metric" style="align-items:flex-start">
      <span class="dash-metric-label" style="padding-top:1px">Cut</span>
      <span class="dash-metric-value" style="text-align:right">
        ${dashFmt(data.billableTons)} Billable tons
        ${data.cutTons !== null ? `<br><span style="color:#64748b;font-size:10px;font-weight:400">Total: ${dashFmt(data.cutTons)} tons / ${fmtLbs(data.cutLbs)}</span>` : ''}
      </span>
    </div>
    <hr class="dash-metric-divider">
    ${combinedRow('Received (Truck)', data.truckTons, truckLbs, data.truckCount, 'trucks')}
    ${combinedRow('Received (Railcar)', data.railTons, railLbs, data.railcarCount, 'railcars')}
    <hr class="dash-metric-divider">
    <div class="dash-metric" style="align-items:flex-start">
      <span class="dash-metric-label" style="padding-top:1px">Net Inv. Change</span>
      <span class="dash-metric-value${netClass}" style="text-align:right">
        ${dashFmt(netChange)} tons${netChangeLbs !== null ? ` / ${fmtLbs(netChangeLbs)}` : ''}
      </span>
    </div>
  </div>`;
}

// Renders 12 skeleton placeholder cards while data loads.
function renderDashSkeletons() {
  const grid = document.getElementById('dashGrid');
  if (!grid) return;
  const MONTHS = ['January','February','March','April','May','June','July','August','September','October','November','December'];
  grid.innerHTML = MONTHS.map(name => `
    <div class="dash-card">
      <div class="dash-card-title">${name}</div>
      <div class="dash-skeleton" style="width:80%"></div>
      <div class="dash-skeleton" style="width:65%"></div>
      <div class="dash-skeleton" style="width:70%"></div>
      <div class="dash-skeleton" style="width:55%"></div>
      <div class="dash-skeleton" style="width:75%"></div>
      <div class="dash-skeleton" style="width:60%"></div>
    </div>`).join('');
}

// State for the dashboard year currently displayed.
let _dashYear = null;
let _dashAvailableYears = [];

// Discovers which years have any history data.
async function discoverDashboardYears() {
  try {
    const [rx, burn] = await Promise.all([
      discoverHistoryMonths('Receiving1'),
      discoverHistoryMonthsForYearlySheet('BurningHistory'),
    ]);
    const yearSet = new Set();
    [...rx, ...burn].forEach(p => yearSet.add(p.year));
    // Always include the current inventory year.
    yearSet.add(getCurrentInventoryPeriod().year);
    return Array.from(yearSet).sort((a, b) => b - a);
  } catch (err) {
    console.warn('Dashboard year discovery failed:', err);
    return [getCurrentInventoryPeriod().year];
  }
}

// Fetches data for all 12 months of the given year and renders the grid.
async function buildDashboardYear(year) {
  _dashYear = year;
  const yearLabel = document.getElementById('dashYearLabel');
  const prevBtn = document.getElementById('dashPrevYear');
  const nextBtn = document.getElementById('dashNextYear');
  if (yearLabel) yearLabel.textContent = year;
  if (prevBtn) prevBtn.disabled = !_dashAvailableYears.includes(year - 1);
  if (nextBtn) nextBtn.disabled = !_dashAvailableYears.includes(year + 1);

  renderDashSkeletons();

  // Fetch all 12 months in parallel.
  const summaries = await Promise.all(
    Array.from({ length: 12 }, (_, m) => fetchMonthSummary(year, m).catch(() => null))
  );

  const grid = document.getElementById('dashGrid');
  if (!grid) return;
  grid.innerHTML = summaries.map((s, m) => renderDashCard(year, m, s)).join('');
}

// One-time init: discovers years, sets default year, wires year nav buttons.
let _dashInitialized = false;
async function initDashboard() {
  if (_dashInitialized) {
    // Already initialized — just refresh to make sure the current period is highlighted.
    buildDashboardYear(_dashYear ?? getCurrentInventoryPeriod().year);
    return;
  }
  _dashInitialized = true;

  _dashAvailableYears = await discoverDashboardYears();
  const defaultYear = _dashAvailableYears[0] ?? getCurrentInventoryPeriod().year;

  document.getElementById('dashPrevYear')?.addEventListener('click', () => {
    if (_dashAvailableYears.includes(_dashYear - 1)) buildDashboardYear(_dashYear - 1);
  });
  document.getElementById('dashNextYear')?.addEventListener('click', () => {
    if (_dashAvailableYears.includes(_dashYear + 1)) buildDashboardYear(_dashYear + 1);
  });
  document.getElementById('dashBackBtn')?.addEventListener('click', showMap);

  buildDashboardYear(defaultYear);
}

function showDashboard() {
  document.getElementById('dashboard').style.display = 'block';
  document.getElementById('map').style.display = 'none';
  document.getElementById('dashToggleBtn').style.display = 'none';
  initDashboard();
}

function showMap() {
  document.getElementById('map').style.display = 'block';
  document.getElementById('dashboard').style.display = 'none';
  document.getElementById('dashToggleBtn').style.display = 'flex';
}

document.getElementById('dashToggleBtn').addEventListener('click', showDashboard);

// Right-click context menu on the dashboard background (outside month cards).
const dashCtxMenu = document.createElement('div');
dashCtxMenu.id = 'dashContextMenu';
dashCtxMenu.style.cssText = 'position:fixed;display:none;z-index:9000;background:#000;border:1px solid #fff;border-radius:4px;color:#fff;';
dashCtxMenu.innerHTML = '<ul style="margin:0;padding:0;list-style:none"><li id="dashReturnItem" style="padding:5px 10px;cursor:pointer">Return to Yardmap</li></ul>';
document.body.appendChild(dashCtxMenu);

document.getElementById('dashReturnItem').addEventListener('click', () => {
  dashCtxMenu.style.display = 'none';
  showMap();
});

document.getElementById('dashboard').addEventListener('contextmenu', e => {
  if (e.target.closest('.dash-card')) return;
  e.preventDefault();
  dashCtxMenu.style.left = e.clientX + 'px';
  dashCtxMenu.style.top = e.clientY + 'px';
  dashCtxMenu.style.display = 'block';
});

document.addEventListener('click', () => { dashCtxMenu.style.display = 'none'; });
document.addEventListener('contextmenu', e => {
  if (!e.target.closest('#dashboard') || e.target.closest('.dash-card')) {
    dashCtxMenu.style.display = 'none';
  }
});
