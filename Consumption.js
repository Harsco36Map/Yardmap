// Bucket Loading (Consumption) — popup data, rendering, and event wiring
// Requires: utils.js loaded first, map.js globals (stockIndexGlobal,
//           POPUP_CONTAINER_STYLE, ACTIVITY_CONTAINER_STYLE, ACTIVITY_TABLE_STYLE,
//           currentMonthBucketPounds, updateMaterialReceivedBanner)

let bucketLoadingCache = { at: 0, data: null };
const bucketHistoryCache = {};
let bucketHistoryMonths = null;
let bucketCurrentPeriod = null;

async function fetchBucketLoadingConsumption(force = false, periodOverride = null) {
  const now = Date.now();
  if (periodOverride) {
    const key = `${periodOverride.year}-${periodOverride.month}`;
    if (!force && bucketHistoryCache[key]) {
      window.currentBucketLoadingRows = bucketHistoryCache[key].rows;
      return bucketHistoryCache[key];
    }
  } else {
    if (!force && bucketLoadingCache.data && (now - bucketLoadingCache.at) < 120000) {
      return bucketLoadingCache.data;
    }
  }

  const toNum = v => {
    if (v == null) return 0;
    if (typeof v === 'number') return Number.isFinite(v) ? v : 0;

    const raw = String(v).trim();
    if (!raw) return 0;

    // Handle common Excel/export formats like "1,234", "(1,234)", or "1,234 lbs".
    const cleaned = raw
      .replace(/,/g, '')
      .replace(/\(([^)]+)\)/, '-$1')
      .replace(/[^0-9.+-]/g, '');

    if (!cleaned || cleaned === '-' || cleaned === '+') return 0;
    const n = Number(cleaned);
    return Number.isFinite(n) ? n : 0;
  };

  const workbook = periodOverride ? await loadHistoryWorkbook(force) : await loadTotalsWorkbook(force);
  const sheet = findSheetByName(workbook, historySheetName('Consumption1', periodOverride));
  if (!sheet) {
    console.warn('Consumption1 sheet not found');
    return null;
  }

  const data = window.XLSX.utils.sheet_to_json(sheet, { header: 1, raw: false });
  if (!Array.isArray(data) || data.length < 2) {
    console.warn('Consumption1 sheet has insufficient data');
    return null;
  }

  const { year: filterYear, month: filterMonth } = periodOverride || getCurrentInventoryPeriod();

  const monthMatches = (date) => {
    return date instanceof Date && !Number.isNaN(date.getTime()) && date.getFullYear() === filterYear && date.getMonth() === filterMonth;
  };

  const toIsoFromDate = (dt) => dt instanceof Date && !Number.isNaN(dt.getTime())
    ? `${dt.getFullYear()}-${String(dt.getMonth() + 1).padStart(2, '0')}-${String(dt.getDate()).padStart(2, '0')}`
    : '';

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
  const poundsCol = findCol('total lbs', 'lbs', 'pounds');
  const tonsCol = findCol('total tons', 'tons');
  const bucketsCol = findCol('buckets loaded', 'buckets');
  const heatsCol = findCol('heats completed', 'heats');

  const resolvedDateCol = dateCol >= 0 ? dateCol : 0;
  const resolvedPoundsCol = poundsCol >= 0 ? poundsCol : 1;
  const resolvedTonsCol = tonsCol >= 0 ? tonsCol : 2;
  const resolvedBucketsCol = bucketsCol >= 0 ? bucketsCol : 3;
  const resolvedHeatsCol = heatsCol >= 0 ? heatsCol : 4;

  const breakdownByIsoDate = {};

  const normalizeKey = (v) => String(v ?? '').trim().toUpperCase();
  const getMaterialForLotOrPile = (lot, pile) => {
    const lotKey = normalizeKey(lot);
    const pileKey = normalizePileCode(pile);

    if (lotKey) {
      const byLot = Object.values(stockIndexGlobal || {}).find(item =>
        normalizeKey(item && item.code) === lotKey
      );
      if (byLot && byLot.material) return String(byLot.material);
    }

    if (pileKey && stockIndexGlobal && stockIndexGlobal[pileKey] && stockIndexGlobal[pileKey].material) {
      return String(stockIndexGlobal[pileKey].material);
    }

    return '';
  };

  const consumption2 = findSheetByName(workbook, historySheetName('Consumption2', periodOverride));
  if (consumption2) {
    const data2 = window.XLSX.utils.sheet_to_json(consumption2, { header: 1, raw: false });
    if (Array.isArray(data2) && data2.length >= 2) {
      const header2 = Array.isArray(data2[0]) ? data2[0] : [];
      const header2Norm = header2.map(h => String(h || '').trim().toLowerCase());
      const findCol2 = (...names) => {
        for (const name of names) {
          const idx = header2Norm.findIndex(h => h === String(name).toLowerCase());
          if (idx >= 0) return idx;
        }
        return -1;
      };

      const date2Col = (() => {
        const i = findCol2('date', 'date material was used');
        return i >= 0 ? i : 0;
      })();
      const pile2Col = (() => {
        const i = findCol2('pile utilized', 'pile used', 'pile', 'pile #');
        return i >= 0 ? i : 1;
      })();
      const lot2Col = (() => {
        const i = findCol2('material lot# utilized', 'material lot # utilized', 'material lot # used', 'material lot #', 'lot #', 'lot');
        return i >= 0 ? i : 2;
      })();
      const pounds2Col = (() => {
        const i = findCol2('pounds utilized', 'total weight of material used in pounds', 'total weight of material used', 'weight (lbs)', 'pounds', 'lbs');
        return i >= 0 ? i : 3;
      })();
      const tons2Col = (() => {
        const i = findCol2('tons utilized', 'tons used', 'tons');
        return i >= 0 ? i : 4;
      })();

      const pileBreakdownByDate = {};

      for (let r2 = 1; r2 < data2.length; r2++) {
        const row2 = data2[r2];
        if (!Array.isArray(row2)) continue;
        const dateCell2 = row2[date2Col];
        const parsed2 = parseXlsxDateCell(dateCell2);

        let iso2 = '';
        if (parsed2) {
          iso2 = toIsoFromDate(parsed2);
        } else {
          const maybeDay2 = Number(String(dateCell2 || '').trim());
          if (Number.isFinite(maybeDay2) && maybeDay2 >= 1 && maybeDay2 <= 31) {
            iso2 = `${filterYear}-${String(filterMonth + 1).padStart(2, '0')}-${String(maybeDay2).padStart(2, '0')}`;
          }
        }

        if (!iso2) continue;

        const pile = normalizePileCode(row2[pile2Col]);
        const lot = String(row2[lot2Col] || '').trim();
        const pounds = toNum(row2[pounds2Col]);
        const tonsCellRaw = row2[tons2Col];
        const tonsCellText = String(tonsCellRaw ?? '').trim();
        const rawTons = toNum(tonsCellRaw);
        const tons = tonsCellText === '' ? (pounds / 2000) : rawTons;

        if (!pile || !Number.isFinite(pounds) || pounds === 0) continue;

        if (!pileBreakdownByDate[iso2]) {
          pileBreakdownByDate[iso2] = {};
        }

        if (!pileBreakdownByDate[iso2][pile]) {
          pileBreakdownByDate[iso2][pile] = { pile, pounds: 0, tons: 0, lots: new Map() };
        }

        pileBreakdownByDate[iso2][pile].pounds += pounds;
        pileBreakdownByDate[iso2][pile].tons += tons;
        if (lot) {
          const material = getMaterialForLotOrPile(lot, pile);
          pileBreakdownByDate[iso2][pile].lots.set(lot, material);
        }
      }

      Object.keys(pileBreakdownByDate).forEach(iso => {
        const pileMap = pileBreakdownByDate[iso];
        breakdownByIsoDate[iso] = Object.values(pileMap)
          .map(item => ({
            pile: item.pile,
            pounds: item.pounds,
            tons: item.tons,
            lotCount: item.lots.size,
            lots: Array.from(item.lots.entries()).map(([lot, material]) => ({ lot, material }))
          }))
          .sort((a, b) => b.pounds - a.pounds || a.pile.localeCompare(b.pile));
      });
    }
  }

  // ---- Consumption4: heat-level material breakdown ----
  const heatBreakdownByIsoDate = {};
  const consumption4 = findSheetByName(workbook, historySheetName('Consumption4', periodOverride));
  if (consumption4) {
    const data4 = window.XLSX.utils.sheet_to_json(consumption4, { header: 1, raw: false });
    if (Array.isArray(data4) && data4.length >= 2) {
      const header4 = Array.isArray(data4[0]) ? data4[0] : [];
      const header4Norm = header4.map(h => String(h || '').trim().toLowerCase());
      const findCol4 = (...names) => {
        for (const name of names) {
          const idx = header4Norm.findIndex(h => h === String(name).toLowerCase());
          if (idx >= 0) return idx;
        }
        return -1;
      };

      const date4Col  = (() => { const i = findCol4('date'); return i >= 0 ? i : 0; })();
      const heat4Col  = (() => { const i = findCol4('heat number', 'heat #', 'heat'); return i >= 0 ? i : 1; })();
      const grade4Col = (() => { const i = findCol4('grade'); return i >= 0 ? i : 2; })();
      const pile4Col  = (() => { const i = findCol4('pile', 'pile #', 'pile utilized'); return i >= 0 ? i : 3; })();
      const lot4Col   = (() => { const i = findCol4('material lot #', 'lot #', 'lot'); return i >= 0 ? i : 4; })();
      const lbs4Col   = (() => { const i = findCol4('total pounds', 'pounds', 'lbs'); return i >= 0 ? i : 5; })();
      const tons4Col  = (() => { const i = findCol4('total tons', 'tons'); return i >= 0 ? i : 6; })();

      // Build per-heat, per-day buckets first, then keep each heat only on its latest day.
      // This prevents carryover heats from showing on both start day and completion day.
      const heatBucketsByNumber = {};

      for (let r4 = 1; r4 < data4.length; r4++) {
        const row4 = data4[r4];
        if (!Array.isArray(row4)) continue;

        const dateCell4 = row4[date4Col];
        const parsed4 = parseXlsxDateCell(dateCell4);
        let iso4 = '';
        if (parsed4) {
          iso4 = toIsoFromDate(parsed4);
        } else {
          const maybeDay4 = Number(String(dateCell4 || '').trim());
          if (Number.isFinite(maybeDay4) && maybeDay4 >= 1 && maybeDay4 <= 31) {
            iso4 = `${filterYear}-${String(filterMonth + 1).padStart(2, '0')}-${String(maybeDay4).padStart(2, '0')}`;
          }
        }
        if (!iso4) continue;

        const heatNum4 = String(row4[heat4Col] || '').trim();
        const grade4   = String(row4[grade4Col] || '').trim();
        const pile4    = normalizePileCode(row4[pile4Col]);
        const lot4     = String(row4[lot4Col] || '').trim();
        const lbs4     = toNum(row4[lbs4Col]);
        const tonsCellRaw4 = row4[tons4Col];
        const tons4    = String(tonsCellRaw4 ?? '').trim() === '' ? (lbs4 / 2000) : toNum(tonsCellRaw4);

        if (!heatNum4) continue;

        if (!heatBucketsByNumber[heatNum4]) {
          heatBucketsByNumber[heatNum4] = {};
        }
        if (!heatBucketsByNumber[heatNum4][iso4]) {
          heatBucketsByNumber[heatNum4][iso4] = { heatNumber: heatNum4, grade: grade4, materials: [] };
        } else if (grade4 && !heatBucketsByNumber[heatNum4][iso4].grade) {
          heatBucketsByNumber[heatNum4][iso4].grade = grade4;
        }

        if (pile4 || lot4 || lbs4) {
          const material4 = getMaterialForLotOrPile(lot4, pile4);
          heatBucketsByNumber[heatNum4][iso4].materials.push({ pile: pile4, lot: lot4, material: material4, pounds: lbs4, tons: tons4 });
        }
      }

      Object.keys(heatBucketsByNumber).forEach(heatNum => {
        const byIso = heatBucketsByNumber[heatNum];
        const isoKeys = Object.keys(byIso);
        if (isoKeys.length === 0) return;

        // ISO date strings sort chronologically (YYYY-MM-DD).
        const latestIso = isoKeys.sort().slice(-1)[0];
        if (!heatBreakdownByIsoDate[latestIso]) heatBreakdownByIsoDate[latestIso] = {};

        const selected = byIso[latestIso];
        if (!heatBreakdownByIsoDate[latestIso][heatNum]) {
          heatBreakdownByIsoDate[latestIso][heatNum] = {
            heatNumber: selected.heatNumber,
            grade: selected.grade,
            materials: []
          };
        }

        heatBreakdownByIsoDate[latestIso][heatNum].materials.push(...(selected.materials || []));
        if (!heatBreakdownByIsoDate[latestIso][heatNum].grade && selected.grade) {
          heatBreakdownByIsoDate[latestIso][heatNum].grade = selected.grade;
        }
      });

      Object.keys(heatBreakdownByIsoDate).forEach(iso => {
        const heatMap = heatBreakdownByIsoDate[iso];
        heatBreakdownByIsoDate[iso] = Object.values(heatMap).sort((a, b) => {
          const na = Number(a.heatNumber), nb = Number(b.heatNumber);
          if (!isNaN(na) && !isNaN(nb)) return na - nb;
          return String(a.heatNumber).localeCompare(String(b.heatNumber));
        });
      });
    }
  }

  const rows = [];
  const allRows = [];
  let totalPounds = 0;
  let totalTons = 0;
  let avgPounds = 0;
  let avgTons = 0;
  let totalRowCount = 0;

  for (let r = 1; r < data.length; r++) {
    const rowData = data[r];
    if (!Array.isArray(rowData)) continue;

    const dateCell = rowData[resolvedDateCol];
    const poundsCell = rowData[resolvedPoundsCol];
    const tonsCell = rowData[resolvedTonsCol];
    const bucketsCell = rowData[resolvedBucketsCol];
    const heatsCell = rowData[resolvedHeatsCol];

    if ((dateCell == null || String(dateCell).trim() === '') && (poundsCell == null || String(poundsCell).trim() === '') && (tonsCell == null || String(tonsCell).trim() === '') && (bucketsCell == null || String(bucketsCell).trim() === '') && (heatsCell == null || String(heatsCell).trim() === '')) {
      continue;
    }

    const dateText = String(dateCell || '').trim();
    const lbs = toNum(poundsCell);
    const tons = toNum(tonsCell);
    const bucketsLoaded = toNum(bucketsCell);
    const heatsCompleted = toNum(heatsCell);
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
    let isoDate = '';

    const parsed = parseXlsxDateCell(dateCell);
    if (monthMatches(parsed)) {
      dayNum = parsed.getDate();
      dateLabel = `${parsed.getMonth() + 1}/${parsed.getDate()}`;
      isoDate = toIsoFromDate(parsed);
    } else {
      const maybeDay = Number(dateText);
      if (Number.isFinite(maybeDay) && maybeDay >= 1 && maybeDay <= 31) {
        dayNum = maybeDay;
        dateLabel = `Day ${maybeDay}`;
        isoDate = `${filterYear}-${String(filterMonth + 1).padStart(2, '0')}-${String(maybeDay).padStart(2, '0')}`;
      } else {
        continue;
      }
    }

    const heatBreakdownForDate = heatBreakdownByIsoDate[isoDate] || [];
    const resolvedHeatsCompleted = heatBreakdownForDate.length > 0
      ? heatBreakdownForDate.length
      : heatsCompleted;

    const rowObj = {
      day: dayNum,
      dateLabel,
      isoDate,
      pounds: lbs,
      tons,
      bucketsLoaded,
      heatsCompleted: resolvedHeatsCompleted,
      breakdown: breakdownByIsoDate[isoDate] || [],
      heatBreakdown: heatBreakdownForDate
    };

    allRows.push(rowObj);
    if (monthMatches(parsed) || !parsed) {
      rows.push(rowObj);
    }
    totalRowCount += 1;
  }

  // If no rows matched the current month filter, fall back to all valid rows.
  if (rows.length === 0 && allRows.length > 0) {
    rows.push(...allRows);
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
      breakdownRowCount: Object.keys(breakdownByIsoDate).length
    },
    totals: {
      totalPounds,
      totalTons,
      avgPounds,
      avgTons,
      totalHeats: rows.reduce((sum, row) => sum + (row.heatsCompleted || 0), 0)
    }
  };

  window.currentBucketLoadingRows = rows;
  if (periodOverride) {
    bucketHistoryCache[`${periodOverride.year}-${periodOverride.month}`] = payload;
  } else {
    bucketLoadingCache = { at: now, data: payload };
    currentMonthBucketPounds = totalPounds;
    updateMaterialReceivedBanner();
  }
  return payload;
}

function renderBucketLoadingPopup(payload) {
  if (!payload) {
    return '<b>Bucket Loading</b><div>No data.</div>';
  }

  const monthLabel = bucketCurrentPeriod
    ? formatMonthYear(bucketCurrentPeriod.year, bucketCurrentPeriod.month)
    : getCurrentInventoryMonthLabel();
  const monthSelectorHtml = buildMonthSelectorHtml('bucket', bucketCurrentPeriod, bucketHistoryMonths);

  const t = payload.totals || {};
  const totalPoundsText = (typeof t.totalPounds === 'number' && isFinite(t.totalPounds)) ? fmtInt(t.totalPounds) : '—';
  const totalTonsText = (typeof t.totalTons === 'number' && isFinite(t.totalTons)) ? fmtTons2(t.totalTons, 2) : '—';
  const avgPoundsText = (typeof t.avgPounds === 'number' && isFinite(t.avgPounds)) ? fmtInt(t.avgPounds) : '—';
  const avgTonsText = (typeof t.avgTons === 'number' && isFinite(t.avgTons)) ? fmtTons2(t.avgTons, 2) : '—';

  const rowsHtml = (payload.rows || []).map(r => `
    <tr class="bucket-loading-main-row" data-date="${esc(r.isoDate)}">
      <td style="padding:2px 6px;white-space:nowrap">${renderDateWithWeekday(r.isoDate, r.dateLabel, { button: true, buttonClass: 'bucket-loading-date-link', dataDate: r.isoDate })}</td>
      <td style="padding:2px 6px;text-align:right">${fmtInt(r.pounds)}</td>
      <td style="padding:2px 6px;text-align:right">${fmtTons2(r.tons, 2)}</td>
      <td style="padding:2px 6px;text-align:right">${fmtInt(r.bucketsLoaded)}</td>
      <td style="padding:2px 6px;text-align:right">${r.heatBreakdown && r.heatBreakdown.length > 0 ? `<button type="button" class="bucket-loading-heat-link" data-date="${esc(r.isoDate)}" style="border:none;background:none;color:#0b57d0;text-decoration:underline;padding:0;cursor:pointer;font:inherit">${fmtInt(r.heatsCompleted)}</button>` : fmtInt(r.heatsCompleted)}</td>
    </tr>
    <tr class="bucket-loading-detail-row" data-date="${esc(r.isoDate)}" style="display:none;background:#fafafa">
      <td colspan="5" style="padding:6px 10px"></td>
    </tr>
    <tr class="bucket-loading-heat-detail-row" data-date="${esc(r.isoDate)}" style="display:none;background:#f0f4ff">
      <td colspan="5" style="padding:6px 10px"></td>
    </tr>
  `).join('');

  const body = `
  <div style="margin-bottom:4px">${monthSelectorHtml}</div>
  <div style="font-weight:700;margin-bottom:6px">Bucket Loading</div>
  <div style="margin-bottom:8px;text-align:center">
    <table style="width:auto;font-size:12px;line-height:1.3;border-collapse:collapse;margin:0 auto;text-align:left">
      <tr><td style="color:#666;padding:2px 6px">Total Consumed</td><td style="text-align:right;padding:2px 6px"><b>${totalPoundsText} lbs</b> <span style="color:#555">(<b>${totalTonsText} tons</b>)</span></td></tr>
      <tr><td style="color:#666;padding:2px 6px">Daily Average</td><td style="text-align:right;padding:2px 6px"><b>${avgPoundsText} lbs</b> <span style="color:#555">(<b>${avgTonsText} tons</b>)</span></td></tr>
    </table>
  </div>
  <div id="bucketLoadingActivity" style="${ACTIVITY_CONTAINER_STYLE}">
    <table style="${ACTIVITY_TABLE_STYLE}">
      <thead>
        <tr style="background:#f2f2f2">
          <th style="text-align:left;padding:2px 6px">Date</th>
          <th style="text-align:right;padding:2px 6px">Pounds</th>
          <th style="text-align:right;padding:2px 6px">Tons</th>
          <th style="text-align:right;padding:2px 6px">Buckets</th>
          <th style="text-align:right;padding:2px 6px">Heats</th>
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
  const monthBtn = container.querySelector('#bucketMonthBtn');
  if (monthBtn) {
    monthBtn.addEventListener('click', async (e) => {
      e.stopPropagation(); e.preventDefault();
      monthBtn.textContent = 'Loading…';
      monthBtn.disabled = true;
      if (!bucketHistoryMonths) bucketHistoryMonths = await discoverHistoryMonths('Consumption1');
      const payload = await fetchBucketLoadingConsumption(false, bucketCurrentPeriod);
      marker.setPopupContent(unescapeAngles(renderBucketLoadingPopup(payload)));
      setTimeout(() => { const el = marker.getPopup()?.getElement(); if (el) wireBucketLoadingPopupEvents(el, marker); }, 0);
    });
  }
  const calToggle = container.querySelector('#bucketMonthCalendarToggle');
  if (calToggle) {
    calToggle.addEventListener('click', (e) => {
      e.preventDefault(); e.stopPropagation();
      const cal = container.querySelector('#bucketMonthCalendar');
      if (cal) cal.style.display = cal.style.display === 'none' ? '' : 'none';
    });
  }
  container.querySelectorAll('.bucketMonthCell').forEach(cell => {
    cell.addEventListener('click', async (e) => {
      e.preventDefault(); e.stopPropagation();
      const val = cell.getAttribute('data-value');
      bucketCurrentPeriod = val === 'current' ? null : { year: +val.split('-')[0], month: +val.split('-')[1] };
      const payload = await fetchBucketLoadingConsumption(false, bucketCurrentPeriod);
      marker.setPopupContent(unescapeAngles(renderBucketLoadingPopup(payload)));
      setTimeout(() => { const el = marker.getPopup()?.getElement(); if (el) wireBucketLoadingPopupEvents(el, marker); }, 0);
    });
  });
  const bucketResetBtn = container.querySelector('#bucketMonthResetBtn');
  if (bucketResetBtn) {
    bucketResetBtn.addEventListener('click', async (e) => {
      e.preventDefault(); e.stopPropagation();
      bucketCurrentPeriod = null;
      const payload = await fetchBucketLoadingConsumption(false, null);
      marker.setPopupContent(unescapeAngles(renderBucketLoadingPopup(payload)));
      setTimeout(() => { const el = marker.getPopup()?.getElement(); if (el) wireBucketLoadingPopupEvents(el, marker); }, 0);
    });
  }

  const toggle = container.querySelector('#bucketLoadingToggle');
  const block = container.querySelector('#bucketLoadingActivity');

  function setBucketLinkState(link, isActive) {
    if (!link) return;
    link.style.color = isActive ? '#c62828' : '#0b57d0';
    link.style.fontWeight = isActive ? '700' : '400';
  }

  function clearBucketLinkStates() {
    container.querySelectorAll('.bucket-loading-date-link, .bucket-loading-heat-link').forEach(link => {
      setBucketLinkState(link, false);
    });
  }

  function setActiveBucketLink(selector, dateKey) {
    clearBucketLinkStates();
    if (!dateKey) return;
    const activeLink = container.querySelector(`${selector}[data-date="${dateKey}"]`);
    if (activeLink) setBucketLinkState(activeLink, true);
  }

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

  // Closes every expanded detail row of either type, optionally skipping one element.
  function closeAllDetailRows(except) {
    container.querySelectorAll('.bucket-loading-detail-row, .bucket-loading-heat-detail-row').forEach(dr => {
      if (dr !== except) dr.style.display = 'none';
    });
    if (!except) clearBucketLinkStates();
  }

  const dateLinks = container.querySelectorAll('.bucket-loading-date-link');
  if (dateLinks && dateLinks.length > 0) {
    dateLinks.forEach(link => {
      link.addEventListener('click', (e) => {
        e.preventDefault();
        e.stopPropagation();

        const dateKey = String(link.getAttribute('data-date') || '').trim();
        if (!dateKey) return;

        const rows = Array.isArray(window.currentBucketLoadingRows) ? window.currentBucketLoadingRows : [];
        const dayRow = rows.find(r => String(r.isoDate || '') === dateKey);
        if (!dayRow) return;

        const detailRow = Array.from(container.querySelectorAll('.bucket-loading-detail-row'))
          .find(dr => dr.getAttribute('data-date') === dateKey);
        if (!detailRow) return;

        const currentlyVisible = detailRow.style.display !== 'none';
        // Close everything (including this row if toggling off)
        closeAllDetailRows(null);
        if (currentlyVisible) return;

        detailRow.style.display = '';
        setActiveBucketLink('.bucket-loading-date-link', dateKey);

        const detailCell = detailRow.querySelector('td');
        if (!detailCell) return;

        const breakdownRows = Array.isArray(dayRow.breakdown) ? dayRow.breakdown : [];
        const breakdownTotal = breakdownRows.reduce((sum, item) => sum + (Number(item.pounds) || 0), 0);
        const breakdownTotalTons = breakdownRows.reduce((sum, item) => sum + (Number(item.tons) || 0), 0);
        const breakdownHtml = breakdownRows.length > 0
          ? breakdownRows.map(item => {
              const lots = Array.isArray(item.lots) ? item.lots : [];
              const lotsText = lots.length > 0
                ? lots.map(x => esc(x.material || '') || '—').join('<br>')
                : '—';
              return `
              <tr>
                <td style="padding:4px 8px;border-bottom:1px solid #eee;white-space:nowrap">${esc(item.pile || '')}</td>
                <td style="padding:4px 8px;border-bottom:1px solid #eee;white-space:nowrap">${lotsText}</td>
                <td style="padding:4px 8px;border-bottom:1px solid #eee;text-align:right">${fmtInt(item.pounds)}</td>
                <td style="padding:4px 8px;border-bottom:1px solid #eee;text-align:right">${fmtTons2(item.tons, 2)}</td>
              </tr>`;
            }).join('')
          : '<tr><td colspan="4" style="padding:8px;color:#666">No pile breakdown data for this day.</td></tr>';

        const totalsHtml = `
          <tr style="background:#f8f8f8">
            <td style="padding:5px 8px;border-top:1px solid #ddd"><b>Daily Total</b></td>
            <td style="padding:5px 8px;border-top:1px solid #ddd">&nbsp;</td>
            <td style="padding:5px 8px;border-top:1px solid #ddd;text-align:right"><b>${fmtInt(breakdownTotal)}</b></td>
            <td style="padding:5px 8px;border-top:1px solid #ddd;text-align:right"><b>${fmtTons2(breakdownTotalTons, 2)}</b></td>
          </tr>`;

        detailCell.innerHTML = `
          <div style="font-size:12px;color:#444;margin-bottom:4px"><b>${esc(dayRow.dateLabel)}</b> pile usage breakdown</div>
          <table style="width:100%;border-collapse:collapse;font-size:12px;background:#fff">
            <thead>
              <tr style="background:#f4f6f8">
                <th style="padding:4px 8px;text-align:left;border-bottom:1px solid #ddd">Pile Utilized</th>
                <th style="padding:4px 8px;text-align:left;border-bottom:1px solid #ddd">Material</th>
                <th style="padding:4px 8px;text-align:right;border-bottom:1px solid #ddd">Pounds</th>
                <th style="padding:4px 8px;text-align:right;border-bottom:1px solid #ddd">Tons</th>
              </tr>
            </thead>
            <tbody>${breakdownHtml}${totalsHtml}</tbody>
          </table>`;
      });
    });
  }

  const heatLinks = container.querySelectorAll('.bucket-loading-heat-link');
  if (heatLinks && heatLinks.length > 0) {
    heatLinks.forEach(link => {
      link.addEventListener('click', (e) => {
        e.preventDefault();
        e.stopPropagation();

        const dateKey = String(link.getAttribute('data-date') || '').trim();
        if (!dateKey) return;

        const rowsData = Array.isArray(window.currentBucketLoadingRows) ? window.currentBucketLoadingRows : [];
        const dayRow = rowsData.find(r => String(r.isoDate || '') === dateKey);
        if (!dayRow) return;

        const heatDetailRow = Array.from(container.querySelectorAll('.bucket-loading-heat-detail-row'))
          .find(dr => dr.getAttribute('data-date') === dateKey);
        if (!heatDetailRow) return;

        const currentlyVisible = heatDetailRow.style.display !== 'none';
        // Close everything (including this row if toggling off)
        closeAllDetailRows(null);
        if (currentlyVisible) return;

        heatDetailRow.style.display = '';
        setActiveBucketLink('.bucket-loading-heat-link', dateKey);

        const detailCell = heatDetailRow.querySelector('td');
        if (!detailCell) return;

        const heats = Array.isArray(dayRow.heatBreakdown) ? dayRow.heatBreakdown : [];
        if (heats.length === 0) {
          detailCell.innerHTML = '<div style="padding:8px;color:#666;font-size:12px">No heat detail data for this day.</div>';
          return;
        }

        // Build accordion: one header row per heat, collapsed material rows below each.
        const accordionRowsHtml = heats.map((heat, hIdx) => {
          const mats = Array.isArray(heat.materials) ? heat.materials : [];
          const heatTotalLbs  = mats.reduce((s, m) => s + (Number(m.pounds) || 0), 0);
          const heatTotalTons = mats.reduce((s, m) => s + (Number(m.tons)   || 0), 0);

          const matsHtml = mats.length > 0
            ? mats.map(m => `
                <tr class="heat-mat-row" data-heat-idx="${hIdx}" style="display:none;background:#f7f8ff">
                  <td style="padding:3px 8px 3px 24px;border-bottom:1px solid #e8edf8;white-space:nowrap">${esc(m.pile || '—')}</td>
                  <td style="padding:3px 8px;border-bottom:1px solid #e8edf8;white-space:normal;overflow-wrap:anywhere;line-height:1.2">
                    <div>${esc(m.material || '—')}</div>
                  </td>
                  <td style="padding:3px 8px;border-bottom:1px solid #e8edf8;text-align:right;white-space:nowrap">${fmtInt(m.pounds)}</td>
                  <td style="padding:3px 8px;border-bottom:1px solid #e8edf8;text-align:right;white-space:nowrap">${fmtTons2(m.tons, 3)}</td>
                </tr>`).join('')
            : `<tr class="heat-mat-row" data-heat-idx="${hIdx}" style="display:none;background:#f7f8ff">
                 <td colspan="4" style="padding:4px 8px 4px 24px;color:#888;font-style:italic">No material records</td>
               </tr>`;

          const chevron = mats.length > 0
            ? `<span class="heat-chevron" style="font-size:10px;margin-right:4px;display:inline-block;transition:transform 0.15s">▶</span>`
            : `<span style="font-size:10px;margin-right:4px;color:#bbb">—</span>`;

          return `
            <tr class="heat-accordion-header" data-heat-idx="${hIdx}"
                style="cursor:${mats.length > 0 ? 'pointer' : 'default'};border-bottom:1px solid #c8d4f0">
              <td style="padding:5px 8px;white-space:nowrap;font-weight:600">${chevron}${esc(heat.heatNumber)}</td>
              <td style="padding:5px 8px;white-space:nowrap">${esc(heat.grade)}</td>
              <td style="padding:5px 8px;text-align:right">${fmtInt(heatTotalLbs)}</td>
              <td style="padding:5px 8px;text-align:right">${fmtTons2(heatTotalTons, 3)}</td>
            </tr>
            <tr class="heat-mat-subheader heat-mat-row" data-heat-idx="${hIdx}" style="display:none;background:#dce4f7;font-size:11px">
              <th style="padding:2px 8px 2px 24px;text-align:left;font-weight:600;border-bottom:1px solid #c8d4f0">Pile</th>
              <th style="padding:2px 8px;text-align:left;font-weight:600;border-bottom:1px solid #c8d4f0">Material</th>
              <th style="padding:2px 8px;text-align:right;font-weight:600;border-bottom:1px solid #c8d4f0">Pounds</th>
              <th style="padding:2px 8px;text-align:right;font-weight:600;border-bottom:1px solid #c8d4f0">Tons</th>
            </tr>
            ${matsHtml}`;
        }).join('');

        detailCell.innerHTML = `
          <div style="font-size:12px;color:#444;margin-bottom:4px">
            <b>${fmtInt(heats.length)} heat${heats.length !== 1 ? 's' : ''} completed on ${esc(dayRow.dateLabel)}</b>
            <span style="color:#888;font-weight:normal"> -- Click a row to expand</span>
          </div>
          <div style="width:100%;max-width:100%;overflow-x:auto">
          <table style="width:100%;table-layout:fixed;border-collapse:collapse;font-size:12px;background:#fff">
            <thead>
              <tr style="background:#e8edf8">
                <th style="padding:4px 8px;text-align:left;border-bottom:1px solid #c8d4f0">Heat #</th>
                <th style="padding:4px 8px;text-align:left;border-bottom:1px solid #c8d4f0">Grade</th>
                <th style="padding:4px 8px;text-align:right;border-bottom:1px solid #c8d4f0">Total Lbs</th>
                <th style="padding:4px 8px;text-align:right;border-bottom:1px solid #c8d4f0">Total Tons</th>
              </tr>
            </thead>
            <tbody>${accordionRowsHtml}</tbody>
          </table>
          </div>`;

        // Wire accordion toggle for each heat header row
        detailCell.querySelectorAll('.heat-accordion-header').forEach(headerRow => {
          const hIdx = headerRow.getAttribute('data-heat-idx');
          const heat = heats[Number(hIdx)];
          if (!heat || !Array.isArray(heat.materials) || heat.materials.length === 0) return;

          headerRow.addEventListener('click', () => {
            const matRows = detailCell.querySelectorAll(`.heat-mat-row[data-heat-idx="${hIdx}"]`);
            const chevron = headerRow.querySelector('.heat-chevron');
            const isOpen = matRows.length > 0 && matRows[0].style.display !== 'none';

            // Collapse all other open heats first
            detailCell.querySelectorAll('.heat-accordion-header').forEach(otherHeader => {
              const oIdx = otherHeader.getAttribute('data-heat-idx');
              if (oIdx === hIdx) return;
              detailCell.querySelectorAll(`.heat-mat-row[data-heat-idx="${oIdx}"]`)
                .forEach(r => { r.style.display = 'none'; });
              const otherChevron = otherHeader.querySelector('.heat-chevron');
              if (otherChevron) otherChevron.style.transform = '';
            });

            matRows.forEach(r => { r.style.display = isOpen ? 'none' : ''; });
            if (chevron) chevron.style.transform = isOpen ? '' : 'rotate(90deg)';
          });
        });
      });
    });
  }
}
