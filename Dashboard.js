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

  const summary = {
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
    buckets:      bucket?.rows ? bucket.rows.reduce((sum, r) => sum + (r.bucketsLoaded || 0), 0) : null,
    brokenTons:    breaking?.totals?.totalTons  ?? null,
    brokenLbs:     breaking?.totals?.totalLbs   ?? null,
    cutTons:       burning?.totals?.netTons     ?? null,
    cutLbs:        burning?.totals?.netLbs      ?? null,
    cuts:          burning?.totals?.cuts        ?? null,
    billableTons:  burning?.totals?.billableTons ?? null,
    billableLbs:   burning?.totals?.billableTons != null ? burning.totals.billableTons * 2000 : null,
  };

  summary.receivedTons = (summary.truckTons !== null || summary.railTons !== null)
    ? (summary.truckTons ?? 0) + (summary.railTons ?? 0)
    : null;
  summary.netChange = (summary.receivedTons !== null || summary.consumedTons !== null)
    ? (summary.receivedTons ?? 0) - (summary.consumedTons ?? 0)
    : null;

  return summary;
}

// Returns true if any field in a summary object has actual data.
function dashSummaryHasData(s) {
  return s && Object.values(s).some(v => v !== null && v !== 0);
}

// Renders the HTML for one month card. data is null for months with no history.
// Metric rows are driven by DASH_METRICS (DashboardCompare.js): each row is a
// button that selects that metric's month-over-month chart, and the month
// title opens the drill-down overlay (DashboardDetail.js).
function renderDashCard(year, month, data) {
  const MONTHS = ['January','February','March','April','May','June','July','August','September','October','November','December'];
  const name = MONTHS[month] || 'Month';
  const { year: curY, month: curM } = getCurrentInventoryPeriod();
  const isCurrent = year === curY && month === curM;
  const cardClass = 'dash-card' + (isCurrent ? ' dash-card--current' : '');

  if (!data || !dashSummaryHasData(data)) {
    return `<div class="${cardClass} dash-card--empty">
      <div style="text-align:center">
        <div class="dash-card-title${isCurrent ? ' current' : ''}" style="border-bottom:none">${name}</div>
        <div style="margin-top:12px">No data</div>
      </div>
    </div>`;
  }

  const rows = DASH_CARD_ROWS.map(key => {
    const cfg = dashMetricByKey(key);
    const v = data[key];
    const valClass = (cfg.signed && v !== null && Number.isFinite(v)) ? (v >= 0 ? ' pos' : ' neg') : '';
    const sign = (cfg.signed && v !== null && Number.isFinite(v) && v >= 0) ? '+' : '';
    const indent = (key === 'truckCount' || key === 'railcarCount') ? ' style="padding-left:16px"' : '';
    return `<button type="button" class="dash-mrow${key === _dashMetric ? ' sel1' : ''}" data-key="${key}" title="View ${cfg.label} month over month">
      <span class="lbl"${indent}>${cfg.label}</span>
      <span class="val${valClass}">${sign}${dashNum(v, cfg.dec)} <span class="unit">${cfg.unit}</span></span>
    </button>` + (DASH_CARD_DIVIDER_AFTER.has(key) ? '<hr class="dash-metric-divider">' : '');
  }).join('');

  return `<div class="${cardClass}">
    <button type="button" class="dash-card-title-btn${isCurrent ? ' current' : ''}" data-month="${month}" title="Open ${name} ${year}">
      <span>${name}${isCurrent ? ' <span style="font-size:9px;vertical-align:middle;opacity:0.8">●</span>' : ''}</span>
      <span class="open-hint">View month &#8599;</span>
    </button>
    ${rows}
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
      discoverHistoryMonths('Receiving'),
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

  // Hand the year's summaries to the month-over-month layer (chips + trend chart).
  _dashSummaries = summaries;
  renderDashCompare();
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
