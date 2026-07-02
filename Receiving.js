// Truck Scales (Receiving) — popup data, rendering, and event wiring
// Requires: utils.js loaded first, map.js globals (stockIndexGlobal, pastDue,
//           POPUP_CONTAINER_STYLE, ACTIVITY_CONTAINER_STYLE, ACTIVITY_TABLE_STYLE,
//           currentMonthTruckWeight, updateMaterialReceivedBanner)

let receivingCache = { at: 0, data: null };
const receivingHistoryCache = {};
let receivingHistoryMonths = null;
let receivingCurrentPeriod = null;

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
    currentMonthTruckWeight = totalWeight;
    updateMaterialReceivedBanner();
  }
  return payload;
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
        const pastDueCodes = new Set((window.pastDue || []).map(p => normalizePileCode(p.code)));
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
                  const isHighlighted = item.isStoppedPile || pastDueCodes.has(normalizePileCode(item.pile || ''));
                  const rowStyle = isHighlighted ? 'background:#ffebee;' : '';
                  const cellStyle = 'padding:2px 6px;' + (isHighlighted ? 'color:#b71c1c;font-weight:700;' : '');
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
