(() => {
  'use strict';

  const data = window.DASHBOARD_DATA;
  const $ = (id) => document.getElementById(id);
  const state = {
    reviewType: 'Weekly',
    periodId: data.meta.currentPeriodId,
    brand: 'all',
    status: 'all',
    search: '',
    page: 1,
    pageSize: 12,
    sortKey: 'rank',
    sortDirection: 'asc'
  };

  const colors = {
    onTime: 'var(--green)',
    late: 'var(--blue)',
    partial: 'var(--orange)',
    missed: 'var(--red)'
  };

  function formatPct(value) {
    return `${Number(value || 0).toFixed(2)}%`;
  }

  function currentPeriod() {
    return data.periods.find((period) => period.id === state.periodId) || data.periods[0];
  }

  function metricsForLocation(location, period = currentPeriod()) {
    const record = period.records[location.name];
    const metrics = record || { onTime: 0, late: 0, partial: 0, missed: 0, hasData: false };
    const completion = Number(metrics.onTime || 0) + Number(metrics.late || 0);
    return { ...location, ...metrics, completion, hasData: Boolean(metrics.hasData) };
  }

  function getStatus(row) {
    if (!row.hasData) return 'No Data';
    if (row.completion >= 90) return 'Excellent';
    if (row.completion >= 75) return 'On Track';
    if (row.completion >= 50) return 'Monitor';
    return 'Needs Attention';
  }

  function allRows(period = currentPeriod()) {
    return data.locations.map((location) => {
      const row = metricsForLocation(location, period);
      return { ...row, status: getStatus(row) };
    });
  }

  function filteredRows() {
    const query = state.search.trim().toLowerCase();
    return allRows()
      .filter((row) => state.brand === 'all' || row.brand === state.brand)
      .filter((row) => state.status === 'all' || row.status === state.status)
      .filter((row) => !query || `${row.name} ${row.brand}`.toLowerCase().includes(query));
  }

  function aggregate(rows) {
    const count = rows.length || 1;
    return rows.reduce((acc, row) => {
      acc.onTime += row.onTime / count;
      acc.late += row.late / count;
      acc.partial += row.partial / count;
      acc.missed += row.missed / count;
      return acc;
    }, { onTime: 0, late: 0, partial: 0, missed: 0, completion: 0, count: rows.length });
  }

  function finalizeAggregate(aggregateValue) {
    aggregateValue.completion = aggregateValue.onTime + aggregateValue.late;
    return aggregateValue;
  }

  function brandRows(rows = allRows()) {
    return data.brands.map((brand) => {
      const subset = rows.filter((row) => row.brand === brand.name);
      return { ...brand, ...finalizeAggregate(aggregate(subset)), locations: subset.length, attention: subset.filter((row) => row.attention).length };
    });
  }

  function updatePeriodOptions() {
    const available = data.periods.filter((period) => period.reviewType === state.reviewType);
    if (!available.some((period) => period.id === state.periodId)) {
      state.periodId = available[available.length - 1].id;
    }
    $('periodFilter').innerHTML = available.map((period) => `<option value="${period.id}">${period.label}</option>`).join('');
    $('periodFilter').value = state.periodId;
  }

  function updateBrandOptions() {
    $('brandFilter').innerHTML = `<option value="all">All Brands</option>${data.brands.map((brand) => `<option value="${escapeHtml(brand.name)}">${escapeHtml(brand.name)}</option>`).join('')}`;
    $('brandFilter').value = state.brand;
  }

  function renderContext() {
    const period = currentPeriod();
    $('contextClient').textContent = data.meta.client;
    $('contextReviewType').textContent = period.reviewType;
    $('contextPeriod').textContent = period.label;
    $('lastRefreshed').textContent = new Date().toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' });
  }

  function renderKpis(summary) {
    const kpis = [
      ['Average Completion', summary.completion, 'primary', '◎'],
      ['Done On-Time', summary.onTime, 'green', '✓'],
      ['Done Late', summary.late, 'blue', '◷'],
      ['Partially Done', summary.partial, 'orange', '◔'],
      ['Missed', summary.missed, 'red', '×'],
      ['Active Locations', summary.count, 'primary', '⌂', false]
    ];
    $('kpiGrid').innerHTML = kpis.map(([label, value, tone, icon, isPercent = true]) => `
      <article class="kpi-card ${tone}">
        <div class="kpi-label"><span>${label}</span><span class="kpi-icon">${icon}</span></div>
        <div class="kpi-value">${isPercent ? formatPct(value) : value}</div>
      </article>
    `).join('');
  }

  function renderSummary(summary, rows) {
    const completion = Math.max(0, Math.min(100, summary.completion));
    const remaining = Math.max(0, 100 - completion);
    $('completionDonut').style.setProperty('--completion', `${completion}%`);
    $('donutValue').textContent = formatPct(completion);
    $('legendCompletion').textContent = formatPct(completion);
    $('legendRemaining').textContent = formatPct(remaining);

    const metrics = [
      ['onTime', 'Done On-Time', summary.onTime, '✓'],
      ['late', 'Done Late', summary.late, '◷'],
      ['partial', 'Partially Done', summary.partial, '◔'],
      ['missed', 'Missed', summary.missed, '×']
    ];
    $('stackedTrack').innerHTML = metrics.map(([key, label, value]) => `<div class="stack-segment ${key}" style="width:${Math.max(0, value)}%" title="${label}: ${formatPct(value)}"></div>`).join('');
    $('breakdownLegend').innerHTML = metrics.map(([key, label, value]) => `<div class="legend-item"><i style="background:${colors[key]}"></i><span>${label}</span><strong>${formatPct(value)}</strong></div>`).join('');
    $('metricList').innerHTML = metrics.map(([key, label, value, icon]) => `<div class="metric-row"><span class="metric-icon" style="color:${colors[key]}">${icon}</span><span>${label}</span><strong style="color:${colors[key]}">${formatPct(value)}</strong></div>`).join('');
    $('locationAverageText').textContent = `Average of ${rows.length} active locations`;
    $('summaryCompletion').textContent = formatPct(summary.completion);
  }

  function renderInsights(rows) {
    const ranked = [...rows].sort((a, b) => b.completion - a.completion);
    const best = ranked[0];
    $('bestLocationName').textContent = best ? best.name : 'No location';
    $('bestLocationMeta').textContent = best ? `${data.meta.client} · ${best.brand}` : 'No data';
    $('bestLocationValue').textContent = best ? formatPct(best.completion) : '0.00%';

    const attention = rows.filter((row) => row.attention || (row.hasData && row.completion < 50)).slice(0, 5);
    $('attentionCount').textContent = `${attention.length} location${attention.length === 1 ? '' : 's'}`;
    $('attentionList').innerHTML = attention.length ? attention.map((row) => `<span>${escapeHtml(row.name)} · ${formatPct(row.completion)}</span>`).join('') : '<span>No locations in the current filtered scope.</span>';

    const period = currentPeriod();
    $('dataStatusText').textContent = `${period.reviewType} report · ${rows.length} locations · Generated from static data · Refreshed ${new Date().toLocaleString()}.`;
  }

  function renderBrandKpis(brands) {
    const sorted = [...brands].sort((a, b) => b.completion - a.completion);
    const best = sorted[0];
    const lowest = [...brands].sort((a, b) => a.completion - b.completion)[0];
    const gap = Math.max(0, (best?.completion || 0) - (lowest?.completion || 0));
    const attention = brands.reduce((sum, brand) => sum + brand.attention, 0);
    const items = [
      ['Brands', brands.length],
      ['Best Brand', best?.name || '—'],
      ['Lowest Brand', lowest?.name || '—'],
      ['Performance Gap', formatPct(gap)],
      ['Attention Locations', attention]
    ];
    $('brandKpis').innerHTML = items.map(([label, value]) => `<div class="brand-kpi"><span>${label}</span><strong>${escapeHtml(String(value))}</strong></div>`).join('');
  }

  function brandInsight(brand) {
    if (brand.name === 'Unassigned Locations') return 'Assign locations';
    if (brand.completion >= 90) return 'On track';
    if (brand.completion > 0) return 'Monitor';
    return 'No completion entered';
  }

  function renderBrands(brands) {
    const sorted = [...brands].sort((a, b) => b.completion - a.completion || a.order - b.order);
    $('brandBars').innerHTML = sorted.map((brand) => `
      <div class="brand-bar-row">
        <button type="button" data-brand-drill="${escapeHtml(brand.name)}">${escapeHtml(brand.name)}</button>
        <div class="brand-track"><div class="brand-fill" style="width:${Math.max(0, brand.completion)}%"></div></div>
        <strong>${formatPct(brand.completion)}</strong>
      </div>
    `).join('');
    $('brandTableBody').innerHTML = sorted.map((brand, index) => `
      <tr data-brand-drill="${escapeHtml(brand.name)}">
        <td>${index + 1}</td><td><strong>${escapeHtml(brand.name)}</strong><br><span class="muted">Client: ${escapeHtml(data.meta.client)}</span></td>
        <td>${brand.locations}</td><td class="number-cell">${formatPct(brand.onTime)}</td><td class="number-cell">${formatPct(brand.late)}</td>
        <td class="number-cell">${formatPct(brand.partial)}</td><td class="number-cell">${formatPct(brand.missed)}</td>
        <td class="number-cell completion-cell">${formatPct(brand.completion)}</td><td>${escapeHtml(brandInsight(brand))}</td>
      </tr>
    `).join('');

    document.querySelectorAll('[data-brand-drill]').forEach((element) => {
      element.addEventListener('click', () => {
        state.brand = element.getAttribute('data-brand-drill');
        state.page = 1;
        $('brandFilter').value = state.brand;
        render();
        document.querySelector('.location-section').scrollIntoView({ behavior: 'smooth' });
      });
    });
  }

  function renderTrend() {
    const periods = data.periods.filter((period) => period.reviewType === state.reviewType);
    const points = periods.map((period) => finalizeAggregate(aggregate(allRows(period))).completion);
    const width = 1100, height = 300, pad = 48;
    const max = Math.max(5, Math.ceil(Math.max(...points) / 5) * 5);
    const x = (index) => pad + (index * (width - pad * 2) / Math.max(1, periods.length - 1));
    const y = (value) => height - pad - (value / max) * (height - pad * 2);
    const path = points.map((value, index) => `${index === 0 ? 'M' : 'L'} ${x(index)} ${y(value)}`).join(' ');
    const area = `${path} L ${x(points.length - 1)} ${height - pad} L ${x(0)} ${height - pad} Z`;
    const grid = [0, .25, .5, .75, 1].map((ratio) => {
      const gy = pad + ratio * (height - pad * 2);
      const label = (max * (1 - ratio)).toFixed(1);
      return `<line x1="${pad}" y1="${gy}" x2="${width - pad}" y2="${gy}" stroke="currentColor" opacity=".10"/><text x="8" y="${gy + 4}" font-size="11" fill="currentColor" opacity=".65">${label}%</text>`;
    }).join('');
    const labels = periods.map((period, index) => `<text x="${x(index)}" y="${height - 14}" text-anchor="middle" font-size="11" fill="currentColor" opacity=".7">${period.shortLabel}</text>`).join('');
    const circles = points.map((value, index) => `<circle cx="${x(index)}" cy="${y(value)}" r="5" fill="#1559b7"/><text x="${x(index)}" y="${y(value) - 13}" text-anchor="middle" font-size="11" font-weight="700" fill="currentColor">${value.toFixed(2)}%</text>`).join('');
    $('trendChart').innerHTML = `<svg viewBox="0 0 ${width} ${height}" role="img"><defs><linearGradient id="areaFill" x1="0" y1="0" x2="0" y2="1"><stop offset="0%" stop-color="#1559b7" stop-opacity=".25"/><stop offset="100%" stop-color="#1559b7" stop-opacity="0"/></linearGradient></defs>${grid}<path d="${area}" fill="url(#areaFill)"/><path d="${path}" fill="none" stroke="#1559b7" stroke-width="4" stroke-linecap="round" stroke-linejoin="round"/>${circles}${labels}</svg>`;
    const currentIndex = periods.findIndex((period) => period.id === state.periodId);
    const currentValue = points[currentIndex] ?? points[points.length - 1];
    const previousValue = currentIndex > 0 ? points[currentIndex - 1] : null;
    $('trendDelta').textContent = previousValue === null ? 'No previous period' : `${currentValue - previousValue >= 0 ? '+' : ''}${(currentValue - previousValue).toFixed(2)} pts`;
  }

  function sortRows(rows) {
    const direction = state.sortDirection === 'asc' ? 1 : -1;
    return [...rows].sort((a, b) => {
      const av = a[state.sortKey];
      const bv = b[state.sortKey];
      if (typeof av === 'string') return av.localeCompare(bv) * direction;
      return ((av || 0) - (bv || 0)) * direction;
    });
  }

  function renderLocationTable(rows) {
    const sorted = sortRows(rows);
    const totalPages = Math.max(1, Math.ceil(sorted.length / state.pageSize));
    state.page = Math.min(state.page, totalPages);
    const start = (state.page - 1) * state.pageSize;
    const pageRows = sorted.slice(start, start + state.pageSize);
    $('locationTableBody').innerHTML = pageRows.map((row) => `
      <tr>
        <td>${row.rank}</td><td><strong>${escapeHtml(row.brand)}</strong></td><td>${escapeHtml(row.name)}</td>
        <td class="number-cell">${formatPct(row.onTime)}</td><td class="number-cell">${formatPct(row.late)}</td>
        <td class="number-cell">${formatPct(row.partial)}</td><td class="number-cell">${formatPct(row.missed)}</td>
        <td class="number-cell completion-cell">${formatPct(row.completion)}</td>
        <td><span class="status-badge status-${statusSlug(row.status)}">${escapeHtml(row.status)}</span></td>
      </tr>
    `).join('');
    const from = sorted.length ? start + 1 : 0;
    const to = Math.min(start + state.pageSize, sorted.length);
    $('paginationInfo').textContent = `Showing ${from}–${to} of ${sorted.length}`;
    $('prevPage').disabled = state.page <= 1;
    $('nextPage').disabled = state.page >= totalPages;
    $('tableSubtitle').textContent = `${sorted.length} filtered locations · ${currentPeriod().label}`;
  }

  function render() {
    const rows = filteredRows();
    const summary = finalizeAggregate(aggregate(rows));
    renderContext();
    renderKpis(summary);
    renderSummary(summary, rows);
    renderInsights(rows);
    const brands = brandRows(allRows()).filter((brand) => state.brand === 'all' || brand.name === state.brand);
    renderBrandKpis(brands);
    renderBrands(brands);
    renderTrend();
    renderLocationTable(rows);
  }

  function exportCsv() {
    const rows = sortRows(filteredRows());
    const headers = ['Rank', 'Client', 'Brand', 'Location', 'Done On-Time', 'Done Late', 'Partially Done', 'Missed', 'Completion', 'Status'];
    const lines = [headers, ...rows.map((row) => [row.rank, data.meta.client, row.brand, row.name, row.onTime.toFixed(2), row.late.toFixed(2), row.partial.toFixed(2), row.missed.toFixed(2), row.completion.toFixed(2), row.status])]
      .map((line) => line.map(csvCell).join(','));
    const blob = new Blob([lines.join('\n')], { type: 'text/csv;charset=utf-8' });
    const url = URL.createObjectURL(blob);
    const link = document.createElement('a');
    link.href = url;
    link.download = `completion-dashboard-${state.periodId}.csv`;
    document.body.appendChild(link);
    link.click();
    link.remove();
    URL.revokeObjectURL(url);
  }

  function csvCell(value) {
    const text = String(value ?? '');
    return `"${text.replaceAll('"', '""')}"`;
  }

  function statusSlug(value) {
    return String(value || '').toLowerCase().replace(/[^a-z0-9]+/g, '-').replace(/^-|-$/g, '');
  }

  function escapeHtml(value) {
    return String(value ?? '').replace(/[&<>'"]/g, (character) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', "'": '&#39;', '"': '&quot;' }[character]));
  }

  function bindEvents() {
    $('reviewTypeFilter').addEventListener('change', (event) => { state.reviewType = event.target.value; updatePeriodOptions(); state.page = 1; render(); });
    $('periodFilter').addEventListener('change', (event) => { state.periodId = event.target.value; state.page = 1; render(); });
    $('brandFilter').addEventListener('change', (event) => { state.brand = event.target.value; state.page = 1; render(); });
    $('statusFilter').addEventListener('change', (event) => { state.status = event.target.value; state.page = 1; render(); });
    $('searchFilter').addEventListener('input', (event) => { state.search = event.target.value; state.page = 1; render(); });
    $('pageSizeSelect').addEventListener('change', (event) => { state.pageSize = Number(event.target.value); state.page = 1; render(); });
    $('prevPage').addEventListener('click', () => { state.page = Math.max(1, state.page - 1); renderLocationTable(filteredRows()); });
    $('nextPage').addEventListener('click', () => { state.page += 1; renderLocationTable(filteredRows()); });
    $('resetButton').addEventListener('click', () => {
      Object.assign(state, { reviewType: 'Weekly', periodId: data.meta.currentPeriodId, brand: 'all', status: 'all', search: '', page: 1, pageSize: 12, sortKey: 'rank', sortDirection: 'asc' });
      $('reviewTypeFilter').value = state.reviewType; updatePeriodOptions(); $('brandFilter').value = 'all'; $('statusFilter').value = 'all'; $('searchFilter').value = ''; $('pageSizeSelect').value = '12'; render();
    });
    $('csvButton').addEventListener('click', exportCsv);
    $('printButton').addEventListener('click', () => window.print());
    $('themeToggle').addEventListener('click', () => document.body.classList.toggle('dark'));
    document.querySelectorAll('th[data-sort]').forEach((header) => {
      header.addEventListener('click', () => {
        const key = header.dataset.sort;
        if (state.sortKey === key) state.sortDirection = state.sortDirection === 'asc' ? 'desc' : 'asc';
        else { state.sortKey = key; state.sortDirection = 'asc'; }
        renderLocationTable(filteredRows());
      });
    });
  }

  updateBrandOptions();
  updatePeriodOptions();
  bindEvents();
  render();
})();
