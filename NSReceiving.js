let railcarCache = { at: 0, data: null };
let railcarHistoryMonths = null;
let railcarCurrentPeriod = null;
const railcarHistoryCache = {};
let railcarPhotoOverlay = null;

async function fetchRailcarSummary(force = false, periodOverride = null) {
  const now = Date.now();
  if (periodOverride) {
    const key = `${periodOverride.year}-${periodOverride.month}`;
    if (!force && railcarHistoryCache[key]) {
      return railcarHistoryCache[key];
    }
  } else {
    if (!force && railcarCache.data && (now - railcarCache.at) < 120000) {
      return railcarCache.data;
    }
  }

  const workbook = periodOverride ? await loadHistoryWorkbook(force) : await loadTotalsWorkbook(force);
  const sheet = findSheetByName(workbook, historySheetName('Railcars', periodOverride));
  if (!sheet) {
    console.warn('Railcars sheet not found');
    return null;
  }

  const data = window.XLSX.utils.sheet_to_json(sheet, { header: 1, raw: false });
  if (!Array.isArray(data) || data.length < 3) {
    console.warn('Railcars sheet has insufficient data');
    return null;
  }

  const railcars = [];
  let i = 2; // skip 2 header rows

  while (i < data.length) {
    const primary = Array.isArray(data[i]) ? data[i] : [];
    const railcarNum = String(primary[0] || '').trim();

    if (!railcarNum) { i++; continue; }

    const next = (i + 1 < data.length && Array.isArray(data[i + 1])) ? data[i + 1] : [];
    const hasSecondary = !String(next[0] || '').trim(); // secondary row has blank col A

    const statusA = String(primary[11] || '').trim();
    const statusB = hasSecondary ? String(next[11] || '').trim() : '';
    const status = [statusA, statusB].filter(Boolean).join(' ');

    railcars.push({
      railcarNum,
      po:           String(primary[1] || '').trim(),
      supplier:     String(primary[2] || '').trim(),
      lot:          String(primary[3] || '').trim(),
      material:     String(primary[4] || '').trim(),
      shipperGross: String(primary[5] || '').trim(),
      shipperTare:  String(primary[6] || '').trim(),
      shipperNet:   String(primary[7] || '').trim(),
      ourGross:     String(primary[8] || '').trim(),
      ourTare:      String(primary[9] || '').trim(),
      pile:         String(primary[10] || '').trim(),
      status,
      mainSupplier: hasSecondary ? String(next[2] || '').trim() : '',
      consistency:  hasSecondary ? String(next[3] || '').trim() : '',
      regrade:      hasSecondary ? String(next[4] || '').trim() : ''
    });

    i += hasSecondary ? 2 : 1;
  }

  const active = railcars.filter(c => !c.status);
  const released = railcars.filter(c => !!c.status);

  const payload = { railcars, active, released };
  if (periodOverride) {
    const key = `${periodOverride.year}-${periodOverride.month}`;
    railcarHistoryCache[key] = payload;
  } else {
    railcarCache = { at: now, data: payload };
    currentMonthRailWeight = released.reduce((sum, car) => {
      const g = parseWeight(car.ourGross);
      const t = parseWeight(car.ourTare);
      return (Number.isFinite(g) && Number.isFinite(t)) ? sum + (g - t) : sum;
    }, 0);
    updateMaterialReceivedBanner();
  }
  return payload;
}

function showRailcarPhoto(railcarNum, monthFolder) {
  if (!railcarPhotoOverlay) {
    const overlay = document.createElement('div');
    overlay.id = 'railcar-photo-overlay';
    overlay.style.cssText = 'position:fixed;top:0;left:0;width:100%;height:100%;background:rgba(0,0,0,0.82);z-index:9999;display:flex;flex-direction:column;align-items:center;justify-content:center;cursor:pointer';
    overlay.innerHTML = '<img id="railcar-photo-img" style="max-width:90vw;max-height:80vh;object-fit:contain;box-shadow:0 4px 32px rgba(0,0,0,0.7);border-radius:4px">'
      + '<div id="railcar-photo-caption" style="color:#fff;margin-top:10px;font-size:14px;font-family:sans-serif;text-align:center"></div>'
      + '<div style="color:#aaa;font-size:12px;margin-top:4px;font-family:sans-serif">Click anywhere or press ESC to close</div>';
    overlay.addEventListener('click', () => { overlay.style.display = 'none'; });
    document.addEventListener('keydown', e => { if (e.key === 'Escape') overlay.style.display = 'none'; });
    document.body.appendChild(overlay);
    railcarPhotoOverlay = overlay;
  }

  const img = railcarPhotoOverlay.querySelector('#railcar-photo-img');
  const caption = railcarPhotoOverlay.querySelector('#railcar-photo-caption');
  img.src = '';
  caption.textContent = 'Loading…';
  railcarPhotoOverlay.style.display = 'flex';

  function tryNext(i) {
    const exts = ['png', 'jpg', 'jpeg'];
    if (i >= exts.length) {
      img.src = '';
      caption.textContent = 'No photo found for ' + railcarNum;
      return;
    }
    const url = 'RailcarPhotos/' + monthFolder + '/' + railcarNum + '.' + exts[i];
    const test = new Image();
    test.onload = () => { img.src = url; caption.textContent = railcarNum + ' — ' + monthFolder; };
    test.onerror = () => tryNext(i + 1);
    test.src = url;
  }
  tryNext(0);
}

function renderRailcarPopup(payload) {
  if (!payload) {
    return '<b>Railroad</b><div>No data.</div>';
  }

  const active = payload.active || [];
  const released = payload.released || [];
  const total = Array.isArray(payload.railcars) ? payload.railcars.length : 0;

  const fmtW = s => {
    const n = parseWeight(s);
    return Number.isFinite(n) ? fmtInt(n) : (s || '—');
  };
  const computeOurNet = car => {
    const g = parseWeight(car.ourGross);
    const t = parseWeight(car.ourTare);
    if (Number.isFinite(g) && Number.isFinite(t)) return fmtInt(g - t);
    return '—';
  };

  const fmtSupplier = car => {
    const line1 = esc(car.supplier || '—');
    if (!car.mainSupplier) return line1;
    return line1 + '<br><span style="font-size:11px;color:#777">' + esc(car.mainSupplier) + '</span>';
  };

  const fmtMaterial = car => {
    const line1 = esc(car.material || '—');
    const extras = [car.regrade, car.consistency].filter(Boolean).map(esc).join(' · ');
    if (!extras) return line1;
    return line1 + '<br><span style="font-size:11px;color:#777">' + extras + '</span>';
  };

  const monthFolder = railcarCurrentPeriod
    ? new Date(railcarCurrentPeriod.year, railcarCurrentPeriod.month, 1).toLocaleString('en-US', { month: 'long' })
    : getCurrentInventoryMonthLabel();
  const monthLabel = railcarCurrentPeriod
    ? formatMonthYear(railcarCurrentPeriod.year, railcarCurrentPeriod.month)
    : getCurrentInventoryMonthLabel();
  const monthSelectorHtml = buildMonthSelectorHtml('railcar', railcarCurrentPeriod, railcarHistoryMonths);

  const activeHtml = active.length > 0
    ? '<table style="width:100%;font-size:12px;border-collapse:collapse">'
      + '<thead><tr style="background:#fff3e0">'
      + '<th style="text-align:left;padding:2px 6px">Car #</th>'
      + '<th style="text-align:left;padding:2px 6px">PO #</th>'
      + '<th style="text-align:left;padding:2px 6px">Supplier</th>'
      + '<th style="text-align:left;padding:2px 6px">Material</th>'
      + '<th style="text-align:right;padding:2px 6px">Shpr Gross</th>'
      + '<th style="text-align:right;padding:2px 6px">Shpr Tare</th>'
      + '<th style="text-align:right;padding:2px 6px">Shpr Net</th>'
      + '<th style="text-align:right;padding:2px 6px">Scale Gross</th>'
      + '</tr></thead><tbody>'
      + active.map(car => {
          const hasScale = !!car.ourGross;
          const rowStyle = 'border-top:1px solid #f5f5f5;' + (hasScale ? 'background:#e8f5e9;' : '');
          const gOurs = parseWeight(car.ourGross);
          const gShpr = parseWeight(car.shipperGross);
          const grossDiff = (Number.isFinite(gOurs) && Number.isFinite(gShpr)) ? (gOurs - gShpr) : null;
          const grossTitle = grossDiff !== null
            ? fmtInt(Math.abs(grossDiff)) + ' lbs ' + (grossDiff > 0 ? "Higher than shipper's" : grossDiff < 0 ? "Lower than shipper's" : 'Exact match')
            : '';
          const scaleGrossCell = hasScale
            ? (grossTitle ? '<span style="text-decoration:underline dotted">' + fmtW(car.ourGross) + '</span>' : fmtW(car.ourGross))
            : '—';
          return '<tr style="' + rowStyle + '">'
            + '<td style="padding:2px 6px;font-weight:600"><button type="button" class="railcar-photo-btn" data-car="' + esc(car.railcarNum) + '" data-month="' + esc(monthFolder) + '" style="background:none;border:none;padding:0;cursor:pointer;font-weight:600;color:#0b57d0;text-decoration:underline;font-size:inherit">' + esc(car.railcarNum) + '</button></td>'
            + '<td style="padding:2px 6px;color:#555">' + esc(car.po || '—') + '</td>'
            + '<td style="padding:2px 6px">' + fmtSupplier(car) + '</td>'
            + '<td style="padding:2px 6px">' + fmtMaterial(car) + '</td>'
            + '<td style="padding:2px 6px;text-align:right">' + fmtW(car.shipperGross) + '</td>'
            + '<td style="padding:2px 6px;text-align:right">' + fmtW(car.shipperTare) + '</td>'
            + '<td style="padding:2px 6px;text-align:right">' + fmtW(car.shipperNet) + '</td>'
            + '<td style="padding:2px 6px;text-align:right"' + (grossTitle ? ' data-offset="' + grossTitle + '"' : '') + '>' + scaleGrossCell + '</td>'
            + '</tr>';
        }).join('')
      + '</tbody></table>'
    : '<div style="font-size:12px;color:#888;padding:6px 0">No active railcars.</div>';

  const releasedHtml = released.length > 0
    ? '<table style="width:100%;font-size:12px;border-collapse:collapse">'
      + '<thead><tr style="background:#e8f5e9">'
      + '<th style="text-align:left;padding:2px 6px">Car #</th>'
      + '<th style="text-align:left;padding:2px 6px">PO #</th>'
      + '<th style="text-align:left;padding:2px 6px">Supplier</th>'
      + '<th style="text-align:left;padding:2px 6px">Material</th>'
      + '<th style="text-align:right;padding:2px 6px">Scale Gross</th>'
      + '<th style="text-align:right;padding:2px 6px">Scale Tare</th>'
      + '<th style="text-align:right;padding:2px 6px">Our Net (lbs)</th>'
      + '<th style="text-align:left;padding:2px 6px">Pile</th>'
      + '<th style="text-align:left;padding:2px 6px">Status</th>'
      + '</tr></thead><tbody>'
      + released.map(car => {
          const gOurs = parseWeight(car.ourGross);
          const gShpr = parseWeight(car.shipperGross);
          const grossDiff = (Number.isFinite(gOurs) && Number.isFinite(gShpr)) ? (gOurs - gShpr) : null;
          const grossTitle = grossDiff !== null
            ? fmtInt(Math.abs(grossDiff)) + ' lbs ' + (grossDiff > 0 ? "Higher than shipper's" : grossDiff < 0 ? "Lower than shipper's" : 'Exact match')
            : '';
          const scaleGrossCell = car.ourGross
            ? (grossTitle ? '<span style="text-decoration:underline dotted">' + fmtW(car.ourGross) + '</span>' : fmtW(car.ourGross))
            : '—';
          const ourNetNum = parseWeight(car.ourGross) - parseWeight(car.ourTare);
          const shprNetNum = parseWeight(car.shipperNet);
          const netDiff = (Number.isFinite(ourNetNum) && Number.isFinite(shprNetNum)) ? (ourNetNum - shprNetNum) : null;
          const netTitle = netDiff !== null
            ? fmtInt(Math.abs(netDiff)) + ' lbs ' + (netDiff > 0 ? "Higher than shipper's" : netDiff < 0 ? "Lower than shipper's" : 'Exact match')
            : '';
          const ourNetCell = netTitle
            ? '<span style="text-decoration:underline dotted">' + computeOurNet(car) + '</span>'
            : computeOurNet(car);
          return '<tr style="border-top:1px solid #f5f5f5">'
            + '<td style="padding:2px 6px;font-weight:600"><button type="button" class="railcar-photo-btn" data-car="' + esc(car.railcarNum) + '" data-month="' + esc(monthFolder) + '" style="background:none;border:none;padding:0;cursor:pointer;font-weight:600;color:#0b57d0;text-decoration:underline;font-size:inherit">' + esc(car.railcarNum) + '</button></td>'
            + '<td style="padding:2px 6px;color:#555">' + esc(car.po || '—') + '</td>'
            + '<td style="padding:2px 6px">' + fmtSupplier(car) + '</td>'
            + '<td style="padding:2px 6px">' + fmtMaterial(car) + '</td>'
            + '<td style="padding:2px 6px;text-align:right"' + (grossTitle ? ' data-offset="' + grossTitle + '"' : '') + '>' + scaleGrossCell + '</td>'
            + '<td style="padding:2px 6px;text-align:right">' + (car.ourTare ? fmtW(car.ourTare) : '—') + '</td>'
            + '<td style="padding:2px 6px;text-align:right"' + (netTitle ? ' data-offset="' + netTitle + '"' : '') + '>' + ourNetCell + '</td>'
            + '<td style="padding:2px 6px">' + esc(car.pile || '—') + '</td>'
            + '<td style="padding:2px 6px">' + esc(car.status || '—') + '</td>'
            + '</tr>';
        }).join('')
      + '</tbody></table>'
    : '<div style="font-size:12px;color:#888;padding:6px 0">No released railcars.</div>';

  const railNetTotal = released.reduce((sum, car) => {
    const g = parseWeight(car.ourGross);
    const t = parseWeight(car.ourTare);
    return (Number.isFinite(g) && Number.isFinite(t)) ? sum + (g - t) : sum;
  }, 0);
  const railNetTotalText = railNetTotal > 0
    ? `<b>${fmtInt(railNetTotal)} lbs</b> <span style="color:#555">(<b>${fmtTons2(railNetTotal / 2000, 2)} tons</b>)</span>`
    : '—';

  const body = `
    <div style="margin-bottom:4px">${monthSelectorHtml}</div>
    <div style="font-weight:700;margin-bottom:6px">Railroad</div>
    <div style="margin-bottom:8px;text-align:center">
      <table style="width:auto;font-size:12px;line-height:1.3;border-collapse:collapse;margin:0 auto;text-align:left">
        <tr><td style="color:#666;padding:2px 6px">Total Railcars</td><td style="text-align:right;padding:2px 6px"><b>${total}</b></td></tr>
        <tr><td style="color:#e65100;padding:2px 6px">&#9679; Active (In Yard)</td><td style="text-align:right;padding:2px 6px"><b>${active.length}</b></td></tr>
        <tr><td style="color:#2e7d32;padding:2px 6px">&#9679; Released</td><td style="text-align:right;padding:2px 6px"><b>${released.length}</b></td></tr>
        <tr><td style="color:#666;padding:2px 6px;padding-top:6px;border-top:1px solid #eee">Total Weight Received</td><td style="text-align:right;padding:2px 6px;padding-top:6px;border-top:1px solid #eee">${railNetTotalText}</td></tr>
      </table>
    </div>
    <div style="display:flex;gap:6px;align-items:center;flex-wrap:wrap">
      ${railcarCurrentPeriod === null ? `<button type="button" id="railcarOnsiteToggle" style="padding:2px 6px;border:1px solid #ddd;border-radius:3px;background:#f8f8f8;cursor:pointer">Railcars Onsite (${active.length})</button>` : ''}
      <button type="button" id="railcarReleasedToggle" style="padding:2px 6px;border:1px solid #ddd;border-radius:3px;background:#f8f8f8;cursor:pointer">Released Railcars (${released.length})</button>
    </div>
    ${railcarCurrentPeriod === null ? `<div id="railcarOnsiteSection" style="display:none;margin-top:6px">
      <div style="font-size:11px;font-weight:700;color:#e65100;margin-bottom:4px;text-transform:uppercase;letter-spacing:0.06em">Railcars Onsite (${active.length})</div>
      <div style="max-height:220px;overflow-y:auto;border:1px solid #ffe0b2;border-radius:3px">
        ${activeHtml}
      </div>
    </div>` : ''}
    <div id="railcarReleasedSection" style="display:none;margin-top:6px">
      <div style="font-size:11px;font-weight:700;color:#2e7d32;margin-bottom:4px;text-transform:uppercase;letter-spacing:0.06em">Released Railcars (${released.length})</div>
      <div style="max-height:260px;overflow-y:auto;border:1px solid #c8e6c9;border-radius:3px">
        ${releasedHtml}
      </div>
    </div>
  `;

  return `<div style="min-width:640px;max-width:100%;width:auto">${body}</div>`;
}

function wireRailcarPopupEvents(container, marker) {
  const railcarMonthBtn = container.querySelector('#railcarMonthBtn');
  if (railcarMonthBtn) {
    railcarMonthBtn.addEventListener('click', async (e) => {
      e.stopPropagation(); e.preventDefault();
      railcarMonthBtn.textContent = 'Loading…';
      railcarMonthBtn.disabled = true;
      if (!railcarHistoryMonths) railcarHistoryMonths = await discoverHistoryMonths('Railcars');
      const payload = await fetchRailcarSummary(false, railcarCurrentPeriod);
      marker.setPopupContent(unescapeAngles(renderRailcarPopup(payload)));
      setTimeout(() => { const el = marker.getPopup()?.getElement(); if (el) wireRailcarPopupEvents(el, marker); }, 0);
    });
  }
  const railcarCalToggle = container.querySelector('#railcarMonthCalendarToggle');
  if (railcarCalToggle) {
    railcarCalToggle.addEventListener('click', (e) => {
      e.preventDefault(); e.stopPropagation();
      const cal = container.querySelector('#railcarMonthCalendar');
      if (cal) cal.style.display = cal.style.display === 'none' ? '' : 'none';
    });
  }
  container.querySelectorAll('.railcarMonthCell').forEach(cell => {
    cell.addEventListener('click', async (e) => {
      e.preventDefault(); e.stopPropagation();
      const val = cell.getAttribute('data-value');
      railcarCurrentPeriod = val === 'current' ? null : { year: +val.split('-')[0], month: +val.split('-')[1] };
      const payload = await fetchRailcarSummary(false, railcarCurrentPeriod);
      marker.setPopupContent(unescapeAngles(renderRailcarPopup(payload)));
      setTimeout(() => { const el = marker.getPopup()?.getElement(); if (el) wireRailcarPopupEvents(el, marker); }, 0);
    });
  });
  const railcarResetBtn = container.querySelector('#railcarMonthResetBtn');
  if (railcarResetBtn) {
    railcarResetBtn.addEventListener('click', async (e) => {
      e.preventDefault(); e.stopPropagation();
      railcarCurrentPeriod = null;
      const payload = await fetchRailcarSummary(false, null);
      marker.setPopupContent(unescapeAngles(renderRailcarPopup(payload)));
      setTimeout(() => { const el = marker.getPopup()?.getElement(); if (el) wireRailcarPopupEvents(el, marker); }, 0);
    });
  }

  const pairs = [
    { btn: container.querySelector('#railcarOnsiteToggle'),  section: container.querySelector('#railcarOnsiteSection') },
    { btn: container.querySelector('#railcarReleasedToggle'), section: container.querySelector('#railcarReleasedSection') }
  ];

  function closePanel(p) {
    if (!p.btn || !p.section) return;
    p.section.style.display = 'none';
    p.btn.textContent = p.btn.textContent.replace(/^Hide\s/, '');
  }

  pairs.forEach(active => {
    if (!active.btn || !active.section) return;
    active.btn.addEventListener('click', () => {
      const isOpen = active.section.style.display !== 'none';
      if (!isOpen) pairs.forEach(other => { if (other !== active) closePanel(other); });
      active.section.style.display = isOpen ? 'none' : '';
      active.btn.textContent = isOpen
        ? active.btn.textContent.replace(/^Hide\s/, '')
        : 'Hide ' + active.btn.textContent;
    });
  });

  container.querySelectorAll('.railcar-photo-btn').forEach(btn => {
    btn.addEventListener('click', e => {
      e.stopPropagation();
      const car = btn.getAttribute('data-car');
      const month = btn.getAttribute('data-month');
      if (car) showRailcarPhoto(car, month);
    });
  });

  // Instant offset tooltip for [data-offset] cells
  const tip = document.createElement('div');
  tip.style.cssText = 'position:fixed;z-index:99999;background:rgba(30,30,30,0.92);color:#fff;font-size:11px;padding:5px 9px;border-radius:3px;pointer-events:none;white-space:nowrap;display:none;box-shadow:0 2px 6px rgba(0,0,0,0.3);';
  document.body.appendChild(tip);

  marker.once('popupclose', () => { if (tip.parentNode) tip.parentNode.removeChild(tip); });

  container.querySelectorAll('[data-offset]').forEach(cell => {
    cell.style.cursor = 'pointer';
    cell.addEventListener('mouseenter', e => {
      tip.textContent = cell.getAttribute('data-offset');
      tip.style.display = 'block';
      tip.style.left = (e.clientX + 14) + 'px';
      tip.style.top  = (e.clientY - 10) + 'px';
    });
    cell.addEventListener('mousemove', e => {
      tip.style.left = (e.clientX + 14) + 'px';
      tip.style.top  = (e.clientY - 10) + 'px';
    });
    cell.addEventListener('mouseleave', () => { tip.style.display = 'none'; });
    cell.addEventListener('click', () => {
      const text = cell.getAttribute('data-offset');
      navigator.clipboard.writeText(text).then(() => {
        const prev = tip.textContent;
        tip.textContent = 'Copied!';
        tip.style.display = 'block';
        setTimeout(() => { tip.textContent = prev; tip.style.display = 'none'; }, 1000);
      });
    });
  });
}
