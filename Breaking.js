// Breaking Pit — popup data, rendering, and event wiring
// Requires: utils.js loaded first, map.js globals (allMarkersData, stockIndexGlobal,
//           POPUP_CONTAINER_STYLE, ACTIVITY_CONTAINER_STYLE, ACTIVITY_TABLE_STYLE)

let breakingCache = { at: 0, data: null };
let breakingHistoryMonths = null;
let breakingCurrentPeriod = null;
const breakingHistoryCache = {};

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

  // Download button
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

  // Download handler for Breaking (current month only)
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
