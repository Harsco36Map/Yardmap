// Burning Station — popup data, rendering, and event wiring
// Requires: utils.js loaded first, map.js globals (allMarkersData, stockIndexGlobal,
//           POPUP_CONTAINER_STYLE, ACTIVITY_CONTAINER_STYLE, ACTIVITY_TABLE_STYLE)

let burningCache = { at: 0, data: null };
let burningHistoryMonths = null;
let burningCurrentPeriod = null;
const burningHistoryCache = {};

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
