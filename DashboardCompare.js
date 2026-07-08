// Historical Dashboard — cross-month metric comparison layer.
// Metric registry, compare chips, 12-month trend charts (SVG, dependency-free),
// table view, and month-card row highlighting.
// Requires: utils.js. Loaded before DashboardDetail.js and Dashboard.js.

// ─── Metric registry ──────────────────────────────────────────────────
// One entry per comparable metric; each key maps onto a fetchMonthSummary()
// field. Adding a metric to the dashboard = adding one row here.
// Cutting always leads with BILLABLE tons; receiving leads with counts.
const DASH_METRICS = [
  { key: 'consumedTons', label: 'Consumed',          unit: 'tons',     dec: 0 },
  { key: 'totalHeats',   label: 'Heats',             unit: 'heats',    dec: 0 },
  { key: 'avgDailyTons', label: 'Avg Daily Use',     unit: 'tons/day', dec: 1, noTotal: true },
  { key: 'buckets',      label: 'Buckets Loaded',    unit: 'buckets',  dec: 0 },
  { key: 'brokenTons',   label: 'Broken',            unit: 'tons',     dec: 0 },
  { key: 'billableTons', label: 'Cut — Billable',    unit: 'tons',     dec: 0 },
  { key: 'cuts',         label: 'Torch Cuts',        unit: 'cuts',     dec: 0 },
  { key: 'receivedTons', label: 'Received (Total)',  unit: 'tons',     dec: 0 },
  { key: 'truckCount',   label: 'Trucks Received',   unit: 'trucks',   dec: 0 },
  { key: 'railcarCount', label: 'Railcars Received', unit: 'cars',     dec: 0 },
  { key: 'truckTons',    label: 'Truck Tons',        unit: 'tons',     dec: 0 },
  { key: 'railTons',     label: 'Rail Tons',         unit: 'tons',     dec: 0 },
  { key: 'netChange',    label: 'Net Inv. Change',   unit: 'tons',     dec: 0, signed: true },
];

function dashMetricByKey(key) {
  return DASH_METRICS.find(m => m.key === key) || null;
}

// Rows shown on each month card, in order; dividers group related metrics.
const DASH_CARD_ROWS = ['consumedTons', 'avgDailyTons', 'totalHeats', 'brokenTons', 'billableTons', 'receivedTons', 'truckCount', 'railcarCount', 'netChange'];
const DASH_CARD_DIVIDER_AFTER = new Set(['totalHeats', 'billableTons', 'railcarCount']);

const DASH_MON3 = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];

// ─── Compare state ─────────────────────────────────────────────────────
let _dashSummaries = [];                 // 12 month summaries for the displayed year (set by buildDashboardYear)
let _dashSelected = ['consumedTons'];    // up to 2 selected metric keys
let _dashTableView = false;

function dashNum(val, dec = 0) {
  return (typeof val === 'number' && isFinite(val))
    ? val.toLocaleString('en-US', { minimumFractionDigits: dec, maximumFractionDigits: dec })
    : '—';
}

// Escapes a string for use inside a double-quoted HTML attribute.
// utils.js esc() covers & < > but not quotes, which the tooltip HTML
// embedded in data-tip attributes needs.
function dashEscAttr(s) {
  return esc(s).replace(/"/g, '&quot;');
}

// ─── SVG bar chart (shared with DashboardDetail.js) ────────────────────
// Thin marks, rounded data-ends, hairline grid, value label on every bar,
// full-height hover hit targets carrying tooltip HTML in data-tip.
function dashRoundedBar(x, y, w, h, r) {
  if (h <= 0.5) return '';
  const rr = Math.min(r, w / 2, h);
  const yb = y + h;
  return `M${x} ${yb} L${x} ${y + rr} Q${x} ${y} ${x + rr} ${y} L${x + w - rr} ${y} Q${x + w} ${y} ${x + w} ${y + rr} L${x + w} ${yb} Z`;
}

function dashBarChart({ values, color, height = 150, dec = 0, unit = '', labelEvery = 1 }) {
  const W = 760, H = height, padL = 6, padR = 6, padT = 18;
  const n = values.length;
  const nums = values.map(v => v.value).filter(v => v !== null && isFinite(v));
  const hasNeg = nums.some(v => v < 0);
  const padB = hasNeg ? 32 : 18;   // room for below-bar value labels + axis row
  const maxV = Math.max(...nums.map(Math.abs), 1);
  const innerW = W - padL - padR, innerH = H - padT - padB;
  const band = innerW / n;
  const barW = Math.min(band - 3, 44);   // keeps a >=2px gap between adjacent bars
  const zeroY = hasNeg ? padT + innerH / 2 : padT + innerH;
  const scale = (hasNeg ? innerH / 2 : innerH) / maxV;
  const lblSize = n > 16 ? 8 : 10;       // dense daily charts get smaller labels

  let grid = '';
  for (let g = 1; g <= 3; g++) {
    const y = padT + innerH * g / 4;
    grid += `<line x1="${padL}" y1="${y}" x2="${W - padR}" y2="${y}" stroke="#293852" stroke-width="1"/>`;
  }

  let bars = '', labels = '', hits = '';
  values.forEach((v, i) => {
    const cx = padL + band * i + (band - barW) / 2;
    if (v.value === null) {
      bars += `<line x1="${cx}" y1="${zeroY - 1}" x2="${cx + barW}" y2="${zeroY - 1}" stroke="#334155" stroke-width="2" stroke-dasharray="3 3"/>`;
    } else {
      const h = Math.abs(v.value) * scale;
      const y = v.value >= 0 ? zeroY - h : zeroY;
      const fill = v.hl ? '#22c55e' : color;
      if (v.value >= 0) bars += `<path d="${dashRoundedBar(cx, y, barW, h, 4)}" fill="${fill}"/>`;
      else bars += `<rect x="${cx}" y="${zeroY}" width="${barW}" height="${h}" rx="4" fill="${fill}"/>`;
      // value label on every bar
      const ly = v.value >= 0 ? y - 4 : zeroY + h + lblSize + 1;
      labels += `<text x="${cx + barW / 2}" y="${ly}" text-anchor="middle" font-size="${lblSize}" fill="#94a3b8" font-family="inherit">${dashNum(v.value, dec)}</text>`;
    }
    if (i % labelEvery === 0) {
      labels += `<text x="${cx + barW / 2}" y="${H - 4}" text-anchor="middle" font-size="10" fill="#64748b" font-family="inherit">${esc(v.axis)}</text>`;
    }
    const tipHtml = v.value === null
      ? `<div class="tip-head">${esc(v.label)}</div><div class="tip-mut">No data</div>`
      : `<div class="tip-head">${esc(v.label)}</div><div><b>${dashNum(v.value, dec)}</b> <span class="tip-mut">${esc(unit)}</span></div>${v.sub ? `<div class="tip-mut">${esc(v.sub)}</div>` : ''}`;
    hits += `<rect class="dash-hit" data-tip="${dashEscAttr(tipHtml)}" x="${padL + band * i}" y="0" width="${band}" height="${H}" fill="transparent"/>`;
  });

  const baseline = `<line x1="${padL}" y1="${zeroY}" x2="${W - padR}" y2="${zeroY}" stroke="#475569" stroke-width="1"/>`;
  return `<div class="dash-chart-box"><svg viewBox="0 0 ${W} ${H}" role="img">${grid}${bars}${baseline}${labels}${hits}</svg></div>`;
}

// ─── Chart tooltip (single element, event delegation) ──────────────────
const _dashTip = document.createElement('div');
_dashTip.id = 'dashTip';
document.body.appendChild(_dashTip);

document.addEventListener('mousemove', e => {
  const hit = e.target.closest ? e.target.closest('.dash-hit') : null;
  if (!hit) {
    if (_dashTip.style.display !== 'none') _dashTip.style.display = 'none';
    return;
  }
  _dashTip.innerHTML = hit.getAttribute('data-tip');
  _dashTip.style.display = 'block';
  const pad = 14, tw = _dashTip.offsetWidth, th = _dashTip.offsetHeight;
  let x = e.clientX + pad, y = e.clientY - th - 10;
  if (x + tw > window.innerWidth - 8) x = e.clientX - tw - pad;
  if (y < 8) y = e.clientY + pad;
  _dashTip.style.left = x + 'px';
  _dashTip.style.top = y + 'px';
});

// ─── Compare rendering ─────────────────────────────────────────────────
function toggleDashMetric(key) {
  if (!dashMetricByKey(key)) return;
  const i = _dashSelected.indexOf(key);
  if (i >= 0) _dashSelected.splice(i, 1);
  else {
    _dashSelected.push(key);
    if (_dashSelected.length > 2) _dashSelected.shift();
  }
  renderDashCompare();
}

function renderDashChips() {
  const chips = document.getElementById('dashChips');
  if (!chips) return;
  chips.innerHTML = DASH_METRICS.map(m => {
    const slot = _dashSelected.indexOf(m.key);
    return `<button type="button" class="dash-chip${slot === 0 ? ' on1' : slot === 1 ? ' on2' : ''}" data-key="${m.key}"><span class="dot"></span>${m.label}</button>`;
  }).join('');
  const clear = document.getElementById('dashClearSel');
  if (clear) clear.style.display = _dashSelected.length ? '' : 'none';
  const hint = document.getElementById('dashCompareHint');
  if (hint) hint.style.display = _dashSelected.length ? 'none' : '';
}

// Values for one metric across the 12 months of the displayed year,
// with month-over-month delta text for the tooltip.
function dashMonthValues(key) {
  const { year: curY, month: curM } = getCurrentInventoryPeriod();
  return _dashSummaries.map((s, i) => {
    const v = (s && s[key] != null && isFinite(s[key])) ? s[key] : null;
    let sub = '';
    if (v !== null) {
      let p = i - 1;
      while (p >= 0 && (!_dashSummaries[p] || _dashSummaries[p][key] == null)) p--;
      if (p >= 0) {
        const pv = _dashSummaries[p][key];
        if (pv) {
          const d = (v - pv) / Math.abs(pv) * 100;
          sub = `${d >= 0 ? '+' : ''}${d.toFixed(1)}% vs ${DASH_MON3[p]}`;
        }
      }
      if (_dashYear === curY && i === curM) sub += (sub ? ' · ' : '') + 'month to date';
    }
    return {
      label: formatMonthYear(_dashYear, i),
      axis: DASH_MON3[i],
      value: v,
      sub,
      hl: _dashYear === curY && i === curM
    };
  });
}

function renderDashTrend() {
  const el = document.getElementById('dashTrend');
  if (!el) return;
  if (!_dashSelected.length || !_dashSummaries.length) {
    el.style.display = 'none';
    el.innerHTML = '';
    return;
  }
  el.style.display = 'block';

  if (_dashTableView) {
    const cols = _dashSelected.map(dashMetricByKey);
    const head = `<tr><th>Month</th>${cols.map(c => `<th>${c.label} (${c.unit})</th><th>&Delta; MoM</th>`).join('')}</tr>`;
    const rows = _dashSummaries.map((s, i) => {
      const { year: curY, month: curM } = getCurrentInventoryPeriod();
      const isCur = _dashYear === curY && i === curM;
      if (!s) return `<tr><td>${DASH_MON3[i]}</td>${cols.map(() => '<td class="dash-delta-nil">—</td><td class="dash-delta-nil">—</td>').join('')}</tr>`;
      return `<tr><td>${DASH_MON3[i]}${isCur ? ' ●' : ''}</td>` + cols.map(c => {
        const v = s[c.key];
        let p = i - 1;
        while (p >= 0 && (!_dashSummaries[p] || _dashSummaries[p][c.key] == null)) p--;
        const pv = (p >= 0 && _dashSummaries[p]) ? _dashSummaries[p][c.key] : null;
        const d = pv ? (v - pv) / Math.abs(pv) * 100 : null;
        const dCls = d === null ? 'dash-delta-nil' : d >= 0 ? 'dash-delta-up' : 'dash-delta-dn';
        return `<td>${dashNum(v, c.dec)}</td><td class="${dCls}">${d === null ? '—' : (d >= 0 ? '+' : '') + d.toFixed(1) + '%'}</td>`;
      }).join('') + '</tr>';
    }).join('');
    el.innerHTML = `<div class="dash-trend-card">
      <div class="dash-trend-head"><span class="dash-trend-title">12-Month Comparison</span>
      <button type="button" class="dash-tv-toggle" id="dashTvBtn">Chart view</button></div>
      <div id="dashTrendTableWrap"><table>${head}${rows}</table></div></div>`;
  } else {
    el.innerHTML = _dashSelected.map((key, slot) => {
      const cfg = dashMetricByKey(key);
      const vals = dashMonthValues(key);
      const nums = vals.filter(v => v.value !== null);
      const color = slot === 0 ? '#3987e5' : '#c98500';
      if (!nums.length) {
        return `<div class="dash-trend-card">
          <div class="dash-trend-head"><span class="dash-trend-title"><span class="swatch" style="background:${color}"></span>${cfg.label} — ${_dashYear}</span></div>
          <div style="color:#64748b;font-size:12px;padding:10px 0">No data for ${_dashYear}.</div></div>`;
      }
      const total = nums.reduce((s, v) => s + v.value, 0);
      const avg = total / nums.length;
      const hi = nums.reduce((a, b) => (b.value > a.value ? b : a), nums[0]);
      const lo = nums.reduce((a, b) => (b.value < a.value ? b : a), nums[0]);
      return `<div class="dash-trend-card">
        <div class="dash-trend-head">
          <span class="dash-trend-title"><span class="swatch" style="background:${color}"></span>${cfg.label} — ${_dashYear} <span style="color:#64748b;font-weight:400;text-transform:none">(${cfg.unit})</span></span>
          <span class="dash-trend-stats">
            ${cfg.noTotal ? '' : `<span>Total <b>${dashNum(total, cfg.dec)}</b></span>`}
            <span>Monthly avg <b>${dashNum(avg, cfg.dec)}</b></span>
            <span>High <b>${hi.axis} · ${dashNum(hi.value, cfg.dec)}</b></span>
            <span>Low <b>${lo.axis} · ${dashNum(lo.value, cfg.dec)}</b></span>
            ${slot === 0 ? '<button type="button" class="dash-tv-toggle" id="dashTvBtn">Table view</button>' : ''}
          </span>
        </div>
        ${dashBarChart({ values: vals, color, dec: cfg.dec, unit: cfg.unit })}
      </div>`;
    }).join('');
  }

  const tv = document.getElementById('dashTvBtn');
  if (tv) tv.addEventListener('click', () => {
    _dashTableView = !_dashTableView;
    renderDashTrend();
  });
}

// Re-applies selection highlight classes to the month-card metric rows
// without rebuilding the grid.
function applyDashCardHighlights() {
  document.querySelectorAll('#dashGrid .dash-mrow').forEach(btn => {
    const slot = _dashSelected.indexOf(btn.getAttribute('data-key'));
    btn.classList.toggle('sel1', slot === 0);
    btn.classList.toggle('sel2', slot === 1);
  });
}

function renderDashCompare() {
  renderDashChips();
  renderDashTrend();
  applyDashCardHighlights();
}

// ─── Event wiring ──────────────────────────────────────────────────────
document.getElementById('dashCompareBar')?.addEventListener('click', e => {
  const chip = e.target.closest('.dash-chip');
  if (!chip) return;
  if (chip.id === 'dashClearSel') {
    _dashSelected = [];
    renderDashCompare();
    return;
  }
  toggleDashMetric(chip.getAttribute('data-key'));
});

document.getElementById('dashGrid')?.addEventListener('click', e => {
  const title = e.target.closest('.dash-card-title-btn');
  if (title) {
    openDashMonth(_dashYear, +title.getAttribute('data-month'));
    return;
  }
  const row = e.target.closest('.dash-mrow');
  if (row) toggleDashMetric(row.getAttribute('data-key'));
});

// Chips are static per registry — render them once at load so the compare bar
// is populated while the first year's data is still loading.
renderDashChips();
