window.addEventListener('error', function (event) {
  setStatus('Script error: ' + (event.message || 'Unknown error'), true);
  const diag = document.getElementById('diagText');
  if (diag) diag.textContent = (event.error && event.error.stack) || event.message || 'Unknown error';
});

window.addEventListener('unhandledrejection', function (event) {
  const reason = event.reason || 'Unknown promise rejection';
  setStatus('Workbook error: ' + (reason.message || reason), true);
  const diag = document.getElementById('diagText');
  if (diag) diag.textContent = String(reason.stack || reason);
});

const state = {
  workbook: null,
  selectedRoundIds: [],
  view: 'points',
  selectedClass: 'all',
  compareSelections: ['', '', ''],
  insightDriverSelection: '',
  insightRivalSelection: '',
  insightRivalSelection2: '',
  driverIndex: []
};

const DEFAULT_WORKBOOK_PATH = 'data/season-results.xlsx';
const DEFAULT_WORKBOOK_NAME = 'season-results.xlsx';


function escapeHtml(value) {
  return String(value ?? '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#039;');
}

function normalizeKey(value) {
  return String(value || '').toLowerCase().replace(/[^a-z0-9]+/g, '').trim();
}

function setStatus(message, isError = false) {
  const el = document.getElementById('status');
  if (!el) return;
  el.textContent = message || '';
  el.classList.toggle('hidden', !message);
  el.style.borderColor = isError ? 'rgba(255,62,62,0.6)' : 'rgba(223,255,0,0.26)';
  el.style.color = isError ? '#ff9a9a' : 'rgba(245,245,245,0.68)';
}

function formatTime(value) {
  if (value === null || value === undefined || value === '') return '—';
  const num = Number(value);
  return Number.isFinite(num) ? num.toFixed(3) : String(value);
}

function formatGap(value) {
  const num = Number(value);
  if (!Number.isFinite(num)) return '—';
  if (Math.abs(num) < 0.0005) return 'LEADER';
  return `+${num.toFixed(3)}`;
}

function roundSortValue(round) {
  return Number(round.roundNumber || String(round.id).match(/\d+/)?.[0] || 0);
}

function driverKey(row) {
  return normalizeKey(`${row.driver}|${row.number || ''}|${row.class || row.cls || ''}`);
}

function labelForDriver(row) {
  const classNumber = [groupClassLabel(row), row.number].filter(Boolean).join(' ');
  return `${row.driver}${classNumber ? ` — ${classNumber}` : ''}`;
}

function groupLabel(row) {
  return row?.displayClass || row?.cls || row?.group || 'UNKNOWN';
}

function baseClassLabel(row) {
  const group = groupLabel(row);
  const base = row?.class || row?.paxClass || row?.baseClass || '';
  if (!base || normalizeKey(base) === normalizeKey(group)) return '';
  return base;
}

function groupClassLabel(row) {
  const group = groupLabel(row);
  const base = baseClassLabel(row);
  return base ? `${group} / ${base}` : group;
}

function driverContextLabel(row, includeNumber = true) {
  return [groupClassLabel(row), includeNumber ? row?.number : '', row?.car].filter(Boolean).join(' · ');
}

function assumptionNote(lines) {
  return `<div class="table-note assumption-note"><strong>Assumptions:</strong> ${lines.map(line => escapeHtml(line)).join(' ')}</div>`;
}

function init() {
  document.getElementById('selectLatest')?.addEventListener('click', () => {
    const latest = [...(state.workbook?.rounds || [])].sort((a, b) => roundSortValue(b) - roundSortValue(a))[0];
    state.selectedRoundIds = (state.workbook?.rounds || []).map(round => round.id);
    afterRoundSelectionChanged();
  });
  document.getElementById('selectAll')?.addEventListener('click', () => {
    state.selectedRoundIds = (state.workbook?.rounds || []).map(round => round.id);
    afterRoundSelectionChanged();
  });
  document.getElementById('clearRounds')?.addEventListener('click', () => {
    state.selectedRoundIds = [];
    afterRoundSelectionChanged();
  });

  autoLoadDefaultWorkbook();
}

document.addEventListener('DOMContentLoaded', init);

async function autoLoadDefaultWorkbook() {
  try {
    setStatus(`Loading ${DEFAULT_WORKBOOK_PATH} from this GitHub Pages site...`);
    const response = await fetch(DEFAULT_WORKBOOK_PATH, { cache: 'no-cache' });
    if (!response.ok) {
      throw new Error(`${DEFAULT_WORKBOOK_PATH} returned HTTP ${response.status}`);
    }
    const buffer = await response.arrayBuffer();
    await handleWorkbookBuffer(buffer, DEFAULT_WORKBOOK_NAME, DEFAULT_WORKBOOK_PATH);
  } catch (error) {
    console.warn('Default workbook auto-load failed:', error);
    setStatus(`Could not auto-load ${DEFAULT_WORKBOOK_PATH}. Add the workbook at that path in the repo and redeploy GitHub Pages.`, true);
    const diag = document.getElementById('diagText');
    if (diag) diag.textContent = String(error?.stack || error);
  }
}

async function handleWorkbookBuffer(buffer, fileName, sourceLabel = fileName) {
  const rawWorkbook = await readXlsxAllSheets(buffer);
  const parsed = parseCompetitionWorkbook(rawWorkbook, fileName);

  if (!parsed.rounds.length) throw new Error('No round tabs with class result data were found.');

  state.workbook = parsed;
  state.selectedRoundIds = parsed.rounds.map(round => round.id);
  state.view = 'points';
  state.selectedClass = 'all';
  state.compareSelections = ['', '', ''];
  state.insightDriverSelection = '';
  state.insightRivalSelection = '';
  state.insightRivalSelection2 = '';

  hydrateMeta();
  renderRoundSelector();
  render();
  updateDiagnostics(rawWorkbook, parsed, sourceLabel);
  setStatus('');
}

function hydrateMeta() {
  const workbook = state.workbook;
  const selectedData = getSelectedData();
  document.getElementById('eventTitle').textContent = workbook?.fileName || 'No data loaded';
  document.getElementById('roundCount').textContent = workbook ? workbook.rounds.length : '—';
  document.getElementById('selectedCount') && (document.getElementById('selectedCount').textContent = state.selectedRoundIds.length || '—');
  document.getElementById('participantCount').textContent = selectedData ? selectedData.summary.length : '—';
  document.getElementById('roundPanel')?.classList.toggle('hidden', !workbook);
}

function renderRoundSelector() {
  const root = document.getElementById('roundSelector');
  if (!root || !state.workbook) return;

  const rounds = [...state.workbook.rounds].sort((a, b) => roundSortValue(b) - roundSortValue(a));
  root.innerHTML = rounds.map(round => `
    <label class="round-option">
      <input type="checkbox" value="${escapeHtml(round.id)}" ${state.selectedRoundIds.includes(round.id) ? 'checked' : ''} />
      <span>
        <strong>${escapeHtml(round.label)}</strong>
        <span>${escapeHtml([round.date, round.location, `${round.entries.length} drivers`].filter(Boolean).join(' · '))}</span>
      </span>
    </label>
  `).join('');

  root.querySelectorAll('input[type="checkbox"]').forEach(input => {
    input.addEventListener('change', () => {
      state.selectedRoundIds = Array.from(root.querySelectorAll('input:checked')).map(node => node.value);
      afterRoundSelectionChanged();
    });
  });
}

function afterRoundSelectionChanged() {
  state.compareSelections = ['', '', ''];
  hydrateMeta();
  renderRoundSelector();
  render();
}

function getSelectedRounds() {
  const ids = new Set(state.selectedRoundIds);
  return (state.workbook?.rounds || []).filter(round => ids.has(round.id));
}

function getSelectedData() {
  const selectedRounds = getSelectedRounds();
  if (!selectedRounds.length) return null;

  const entries = selectedRounds.flatMap(round => round.entries.map(entry => ({ ...entry, round })));
  const classes = {};
  const classOrder = [];

  entries.forEach(entry => {
    const cls = entry.displayClass || entry.class || entry.cls || 'UNKNOWN';
    if (!classes[cls]) {
      classes[cls] = [];
      classOrder.push(cls);
    }
    classes[cls].push(entry);
  });

  Object.values(classes).forEach(rows => {
    rows.sort((a, b) => {
      if (selectedRounds.length === 1) return Number(a.position || 9999) - Number(b.position || 9999);
      const roundDelta = roundSortValue(b.round) - roundSortValue(a.round);
      if (roundDelta) return roundDelta;
      return Number(a.position || 9999) - Number(b.position || 9999);
    });
  });

  const overall = entries
    .filter(row => Number.isFinite(Number(row.bestRaw)))
    .slice()
    .sort((a, b) => Number(a.bestRaw) - Number(b.bestRaw))
    .map((row, index) => ({ ...row, rank: index + 1, time: row.bestRaw, overallRank: index + 1 }));

  const pax = entries
    .filter(row => Number.isFinite(Number(row.bestPax)))
    .slice()
    .sort((a, b) => Number(a.bestPax) - Number(b.bestPax))
    .map((row, index) => ({ ...row, rank: index + 1, time: row.bestPax, paxRank: index + 1 }));

  const rankLookup = new Map();
  overall.forEach(row => rankLookup.set(driverKey(row) + '|' + row.round.id, { overallRank: row.rank }));
  pax.forEach(row => {
    const key = driverKey(row) + '|' + row.round.id;
    rankLookup.set(key, { ...(rankLookup.get(key) || {}), paxRank: row.rank });
  });
  entries.forEach(row => Object.assign(row, rankLookup.get(driverKey(row) + '|' + row.round.id) || {}));

  const summary = buildSummary(entries);
  state.driverIndex = summary.map(row => ({ ...row, label: labelForDriver(row) }));

  return { selectedRounds, entries, classes, classOrder: classOrder.sort(), overall, pax, summary };
}

function buildSummary(entries) {
  const map = new Map();
  entries.forEach(row => {
    const key = driverKey(row);
    const existing = map.get(key) || {
      driver: row.driver,
      number: row.number,
      car: row.car,
      cls: row.cls,
      class: row.class,
      displayClass: row.displayClass,
      rows: [],
      roundLabels: []
    };
    existing.rows.push(row);
    existing.roundLabels.push(row.round.label);
    existing.car = existing.car || row.car;
    map.set(key, existing);
  });

  return Array.from(map.values()).map(item => {
    const raws = item.rows.map(row => Number(row.bestRaw)).filter(Number.isFinite);
    const paxes = item.rows.map(row => Number(row.bestPax)).filter(Number.isFinite);
    const positions = item.rows.map(row => Number(row.position)).filter(Number.isFinite);
    const bestRaw = raws.length ? Math.min(...raws) : null;
    const bestPax = paxes.length ? Math.min(...paxes) : null;
    const avgRaw = raws.length ? raws.reduce((a, b) => a + b, 0) / raws.length : null;
    const avgPax = paxes.length ? paxes.reduce((a, b) => a + b, 0) / paxes.length : null;
    const bestClassPosition = positions.length ? Math.min(...positions) : null;
    const classWins = positions.filter(pos => pos === 1).length;

    return {
      ...item,
      bestRaw,
      bestPax,
      avgRaw,
      avgPax,
      bestClassPosition,
      classWins,
      roundsCount: item.rows.length,
      cleanRuns: item.rows.flatMap(row => row.runObjects || []).filter(run => run.isClean).length,
      dnfRuns: item.rows.flatMap(row => row.runObjects || []).filter(run => run.isDnf).length,
      rerunRuns: item.rows.flatMap(row => row.runObjects || []).filter(run => run.isRerun).length
    };
  }).sort((a, b) => {
    const paxDelta = Number(a.bestPax ?? 9999) - Number(b.bestPax ?? 9999);
    if (paxDelta) return paxDelta;
    return Number(a.bestRaw ?? 9999) - Number(b.bestRaw ?? 9999);
  }).map((item, index) => ({ ...item, rank: index + 1 }));
}

function setView(view) {
  state.view = view;
  render();
}

function getViewTitle() {
  if (state.view === 'points') return 'POINTS';
  if (state.view === 'insights') return 'INSIGHTS';
  return 'POINTS';
}

function renderViewDock(data) {
  return `
    <div class="view-dock focused-tool-nav">
      <button class="view-dock-trigger" type="button" data-view-dock-toggle aria-label="Toggle view controls" aria-expanded="false">
        <span class="view-dock-bars" aria-hidden="true"><span></span><span></span><span></span></span>
      </button>
      <div class="view-dock-panel">
        <div class="view-dock-kicker">TOOL SECTIONS</div>
        ${['points', 'insights'].map(view => `
          <button data-view-button="${view}" class="view-button ${state.view === view ? 'active' : ''}" type="button">${view === 'points' ? 'POINTS BY CLASS' : 'DRIVER INSIGHTS'}</button>
        `).join('')}
      </div>
    </div>
  `;
}

function render() {
  const root = document.getElementById('rankings');
  if (!root) return;

  const data = getSelectedData();
  hydrateMeta();

  if (!state.workbook) {
    root.innerHTML = '';
    return;
  }
  if (!data) {
    state.selectedRoundIds = (state.workbook?.rounds || []).map(round => round.id);
  }

  if (!['points', 'insights'].includes(state.view)) state.view = 'points';
  if (state.view === 'points') renderSeasonPoints(data);
  if (state.view === 'insights') renderDriverInsights(data);

  attachGlobalHandlers();
}

function selectedRoundLabel(data) {
  if (!data.selectedRounds.length) return 'No rounds selected';
  if (data.selectedRounds.length === 1) {
    const r = data.selectedRounds[0];
    return [r.label, r.date, r.location].filter(Boolean).join(' · ');
  }
  const sorted = [...data.selectedRounds].sort((a, b) => roundSortValue(a) - roundSortValue(b));
  return `${sorted[0].label} through ${sorted[sorted.length - 1].label} · ${data.selectedRounds.length} rounds selected`;
}

function renderSummary(data) {
  const root = document.getElementById('rankings');

  if ((data.selectedRounds || []).length > 1) {
    const sections = sortedSelectedRounds(data).map(round => {
      const entries = entriesForRound(data, round).map(entry => ({ ...entry, round }));
      const summaryRows = buildSummary(entries);
      const classCount = new Set(entries.map(row => row.displayClass || row.class || row.cls || 'UNKNOWN')).size;
      return `
        <section class="round-class-group" data-round="${escapeHtml(round.id)}">
          ${renderRoundBreakoutHeader(round, `${summaryRows.length} DRIVER${summaryRows.length === 1 ? '' : 'S'}`)}
          <div class="card-body">
            ${renderSummaryMetrics(entries, summaryRows, 1, classCount)}
            ${renderSummaryTable(summaryRows)}
          </div>
        </section>
      `;
    }).join('');

    root.innerHTML = `
      <section class="results-shell">
        ${renderViewDock(data)}
        <section class="card results-card">
          <div class="card-header">
            <div class="class-title">
              <div class="acr-tag">${escapeHtml(getViewTitle())}</div>
              <div class="header-main">Competition Summary</div>
              <div class="driver-sub">${escapeHtml(selectedRoundLabel(data))}</div>
            </div>
            <div class="class-count">${data.summary.length} UNIQUE DRIVER${data.summary.length === 1 ? '' : 'S'}</div>
          </div>
        </section>
        ${sections || renderInlineEmpty('No summary rows found for the selected rounds.')}
      </section>
    `;
    return;
  }

  const classCount = data.classOrder.length;
  root.innerHTML = `
    <section class="results-shell">
      ${renderViewDock(data)}
      <section class="card results-card">
        <div class="card-header">
          <div class="class-title">
            <div class="acr-tag">${escapeHtml(getViewTitle())}</div>
            <div class="header-main">Competition Summary</div>
            <div class="driver-sub">${escapeHtml(selectedRoundLabel(data))}</div>
          </div>
          <div class="class-count">${data.summary.length} DRIVER${data.summary.length === 1 ? '' : 'S'}</div>
        </div>
        <div class="card-body">
          ${renderSummaryMetrics(data.entries, data.summary, data.selectedRounds.length, classCount)}
          ${renderSummaryTable(data.summary)}
        </div>
      </section>
    </section>
  `;
}

function metricCard(label, value) {
  return `<div class="summary-card"><div class="metric-label">${escapeHtml(label)}</div><div class="metric-value">${escapeHtml(value)}</div></div>`;
}

function renderSimpleResults(data, title, rows, timeLabel) {
  const root = document.getElementById('rankings');

  if ((data.selectedRounds || []).length > 1) {
    const isPax = /pax/i.test(timeLabel) || /pax/i.test(title);
    const sections = sortedSelectedRounds(data).map(round => {
      const roundEntries = entriesForRound(data, round);
      const roundRows = isPax ? rankPaxRows(roundEntries) : rankOverallRows(roundEntries);
      return `
        <section class="round-class-group" data-round="${escapeHtml(round.id)}">
          ${renderRoundBreakoutHeader(round, `${roundRows.length} SOURCE ROW${roundRows.length === 1 ? '' : 'S'}`)}
          ${renderPodium(roundRows, timeLabel)}
          <section class="card results-card">
            <div class="card-header">
              <div class="class-title">
                <div class="acr-tag">${escapeHtml(getViewTitle())}</div>
                <div class="header-main">${escapeHtml(title)}</div>
                <div class="driver-sub">${escapeHtml([round.label, round.date, round.location].filter(Boolean).join(' · '))}</div>
              </div>
              <div class="class-count">${roundRows.length} SOURCE ROW${roundRows.length === 1 ? '' : 'S'}</div>
            </div>
            <div class="card-body">
              ${roundRows.map(row => renderResultRow(row, timeLabel, row.time)).join('') || emptyRow('No event rows found')}
            </div>
          </section>
        </section>
      `;
    }).join('');

    root.innerHTML = `
      <section class="results-shell">
        ${renderViewDock(data)}
        <section class="card results-card">
          <div class="card-header">
            <div class="class-title">
              <div class="acr-tag">${escapeHtml(getViewTitle())}</div>
              <div class="header-main">${escapeHtml(title)}</div>
              <div class="driver-sub">${escapeHtml(selectedRoundLabel(data))}</div>
            </div>
            <div class="class-count">EVENT SECTIONS</div>
          </div>
        </section>
        ${sections || renderInlineEmpty('No event rows found for the selected rounds.')}
      </section>
    `;
    return;
  }

  root.innerHTML = `
    ${renderPodium(rows, timeLabel)}
    <section class="results-shell">
      ${renderViewDock(data)}
      <section class="card results-card">
        <div class="card-header">
          <div class="class-title">
            <div class="acr-tag">${escapeHtml(getViewTitle())}</div>
            <div class="header-main">${escapeHtml(title)}</div>
            <div class="driver-sub">${escapeHtml(selectedRoundLabel(data))}</div>
          </div>
          <div class="class-count">${rows.length} SOURCE ROW${rows.length === 1 ? '' : 'S'}</div>
        </div>
        <div class="card-body">
          ${rows.map(row => renderResultRow(row, timeLabel, row.time)).join('') || emptyRow('No event rows found')}
        </div>
      </section>
    </section>
  `;
}

function renderPodium(rows, label) {
  const top = rows.slice(0, 3);
  if (!top.length) return '';
  return `
    <section class="podium">
      ${top.map((row, index) => `
        <article class="podium-card">
          <div class="podium-rank">P${escapeHtml(row.rank || index + 1)}</div>
          <div class="podium-name">${escapeHtml(row.driver)}</div>
          <div class="podium-sub">${escapeHtml([row.round?.label, row.displayClass || row.class, row.number, row.car].filter(Boolean).join(' · '))}</div>
          <div class="podium-time">${escapeHtml(formatTime(row.time))}</div>
          <div class="time-label">${escapeHtml(label)}</div>
        </article>
      `).join('')}
    </section>
  `;
}

function renderClassResults(data) {
  const root = document.getElementById('rankings');
  const visible = state.selectedClass === 'all' ? data.classOrder : [state.selectedClass];
  const selectedRounds = data.selectedRounds || [];

  if (!visible.length) {
    root.innerHTML = renderEmptyCard('CLASS', 'No class data found');
    return;
  }

  const multiRound = selectedRounds.length > 1;
  const content = multiRound
    ? selectedRounds
        .slice()
        .sort((a, b) => roundSortValue(b) - roundSortValue(a))
        .map(round => renderRoundClassGroup(data, round, visible))
        .join('')
    : visible.map((cls, index) => renderClassCard(data, cls, data.classes?.[cls] || [], index)).join('');

  root.innerHTML = `<section class="results-shell">${renderViewDock(data)}${content || renderInlineEmpty('No class rows found for the selected round filters.')}</section>`;
}

function renderRoundClassGroup(data, round, visibleClasses) {
  const cards = visibleClasses
    .map(cls => {
      const rows = (data.classes?.[cls] || [])
        .filter(row => row.round?.id === round.id)
        .slice()
        .sort((a, b) => Number(a.position || 9999) - Number(b.position || 9999));
      if (!rows.length) return '';
      return renderClassCard(data, cls, rows, 1, round);
    })
    .join('');

  if (!cards) return '';

  return `
    <section class="round-class-group" data-round="${escapeHtml(round.id)}">
      <div class="round-group-header">
        <div>
          <div class="round-kicker">EVENT BREAKOUT</div>
          <div class="round-group-title">${escapeHtml(round.label || round.name || round.id)}</div>
          <div class="driver-sub">${escapeHtml([round.date, round.location].filter(Boolean).join(' · '))}</div>
        </div>
        <div class="class-count">${round.entries?.length || 0} SOURCE ROW${(round.entries?.length || 0) === 1 ? '' : 'S'}</div>
      </div>
      <div class="round-class-grid">${cards}</div>
    </section>
  `;
}

function renderClassCard(data, cls, rows, index = 0, round = null) {
  const roundLabel = round ? (round.label || round.name || round.id) : selectedRoundLabel(data);
  const card = `
    <section class="card results-card" data-class="${escapeHtml(cls)}">
      <div class="card-header">
        <div class="class-title">
          <div class="acr-tag">${escapeHtml(cls)}</div>
          <div class="header-main">Class Results</div>
          <div class="driver-sub">${escapeHtml(roundLabel)}</div>
        </div>
        <div class="class-count">${rows.length} DRIVER${rows.length === 1 ? '' : 'S'}</div>
      </div>
      <div class="card-body">
        ${rows.map(row => renderResultRow(row, 'Best Raw / PAX ' + formatTime(row.bestPax), row.bestRaw)).join('') || emptyRow('No class rows found')}
      </div>
    </section>
  `;

  if (index === 0) return `<section class="results-shell">${renderViewDock(data)}${card}</section>`;
  return card;
}

function renderInlineEmpty(message) {
  return `<section class="card results-card"><div class="card-body">${emptyRow(message)}</div></section>`;
}


function placePoints(place) {
  const pos = Number(place);
  if (!Number.isFinite(pos) || pos < 1) return 0;
  if (pos === 1) return 200;
  if (pos === 2) return 160;
  if (pos === 3) return 130;
  if (pos === 4) return 110;
  if (pos === 5) return 90;
  if (pos === 6) return 70;
  if (pos === 7) return 50;
  if (pos === 8) return 30;
  if (pos === 9) return 20;
  return Math.max(1, 29 - pos); // 10th = 19, 11th = 18, down to minimum 1
}

function workbookSeasonYear() {
  const workbook = state.workbook;
  const searchText = [
    workbook?.seasonYear,
    workbook?.fileName,
    ...(workbook?.rounds || []).flatMap(round => [round.title, round.date, round.label, round.sheetName])
  ].filter(Boolean).join(' ');
  const match = String(searchText).match(/(20\d{2})/);
  return match ? match[1] : '';
}

function seasonTitle() {
  const year = workbookSeasonYear();
  return year ? `${year} Season` : 'Season';
}

function renderSeasonPoints(data) {
  const root = document.getElementById('rankings');
  const points = buildSeasonPoints();
  const allRounds = state.workbook?.rounds || [];

  if (!points.classOrder.length) {
    root.innerHTML = renderEmptyCard('POINTS', 'No season points could be calculated from the workbook.');
    return;
  }

  const visibleClasses = state.selectedClass === 'all' ? points.classOrder : points.classOrder.filter(cls => cls === state.selectedClass);
  const cards = visibleClasses.map(cls => renderPointsClassCard(cls, points.classes[cls], points)).join('');

  root.innerHTML = `
    <section class="results-shell">
      ${renderViewDock(data)}
      <section class="card results-card">
        <div class="card-header">
          <div class="class-title">
            <div class="acr-tag">POINTS</div>
            <div class="header-main">${escapeHtml(seasonTitle())} Points Standings</div>
          </div>
          <div class="class-count">${allRounds.length} EVENT${allRounds.length === 1 ? '' : 'S'} · ${points.dropCount} DROP${points.dropCount === 1 ? '' : 'S'}</div>
        </div>
        <div class="card-body">
          <div class="summary-grid">
            ${metricCard('Events Scored', allRounds.length)}
            ${metricCard('Drop Rule', `1 per 4 events`)}
            ${metricCard('Drops Applied', points.dropCount)}
            ${metricCard('Classes', points.classOrder.length)}
          </div>
          <label class="points-class-filter"><span>Class Filter</span><select data-class-filter><option value="all" ${state.selectedClass === 'all' ? 'selected' : ''}>ALL CLASSES</option>${points.classOrder.map(cls => `<option value="${escapeHtml(cls)}" ${state.selectedClass === cls ? 'selected' : ''}>${escapeHtml(cls)}</option>`).join('')}</select></label>
          <div class="driver-sub">Points: 1st 200, 2nd 160, 3rd 130, 4th 110, 5th 90, 6th 70, 7th 50, 8th 30, 9th 20, 10th 19, then decreasing by 1 down to 1. Missed events count as 0 before drops are applied.</div>
        </div>
      </section>
      ${cards}
    </section>
  `;
}

function buildSeasonPoints() {
  const rounds = state.workbook?.rounds || [];
  const dropCount = Math.floor(rounds.length / 4);
  const classes = {};
  const classOrder = [];

  rounds.forEach(round => {
    (round.entries || []).forEach(entry => {
      const cls = entry.displayClass || entry.class || entry.cls || 'UNKNOWN';
      if (!classes[cls]) {
        classes[cls] = new Map();
        classOrder.push(cls);
      }
      const key = seasonDriverKey(entry);
      const existing = classes[cls].get(key) || {
        driver: entry.driver,
        number: entry.number,
        car: entry.car,
        displayClass: cls,
        rows: [],
        pointsByRound: new Map()
      };
      existing.number = existing.number || entry.number;
      existing.car = existing.car || entry.car;
      existing.rows.push(entry);
      existing.pointsByRound.set(round.id, {
        round,
        place: entry.position,
        points: placePoints(entry.position),
        bestRaw: entry.bestRaw,
        bestPax: entry.bestPax
      });
      classes[cls].set(key, existing);
    });
  });

  const built = {};
  classOrder.sort().forEach(cls => {
    const rows = Array.from(classes[cls].values()).map(driver => {
      const eventScores = rounds.map(round => driver.pointsByRound.get(round.id) || {
        round,
        place: null,
        points: 0,
        missed: true
      });
      const sortedForDrops = eventScores
        .map((score, index) => ({ ...score, originalIndex: index }))
        .sort((a, b) => Number(a.points) - Number(b.points));
      const droppedIndexes = new Set(sortedForDrops.slice(0, dropCount).map(score => score.originalIndex));
      const grossPoints = eventScores.reduce((sum, score) => sum + Number(score.points || 0), 0);
      const droppedPoints = eventScores.reduce((sum, score, index) => sum + (droppedIndexes.has(index) ? Number(score.points || 0) : 0), 0);
      const netPoints = grossPoints - droppedPoints;
      const wins = eventScores.filter(score => Number(score.place) === 1).length;
      const podiums = eventScores.filter(score => Number(score.place) >= 1 && Number(score.place) <= 3).length;
      const eventsRun = eventScores.filter(score => !score.missed).length;
      const bestFinishes = eventScores.map(score => Number(score.place)).filter(Number.isFinite);
      const bestFinish = bestFinishes.length ? Math.min(...bestFinishes) : null;

      return {
        ...driver,
        eventScores,
        droppedIndexes,
        grossPoints,
        droppedPoints,
        netPoints,
        wins,
        podiums,
        eventsRun,
        bestFinish
      };
    }).sort((a, b) => {
      const netDelta = Number(b.netPoints) - Number(a.netPoints);
      if (netDelta) return netDelta;
      const winDelta = Number(b.wins) - Number(a.wins);
      if (winDelta) return winDelta;
      const podiumDelta = Number(b.podiums) - Number(a.podiums);
      if (podiumDelta) return podiumDelta;
      const bestDelta = Number(a.bestFinish ?? 9999) - Number(b.bestFinish ?? 9999);
      if (bestDelta) return bestDelta;
      return String(a.driver).localeCompare(String(b.driver));
    }).map((row, index) => ({ ...row, rank: index + 1 }));
    built[cls] = rows;
  });

  return { rounds, dropCount, classOrder: classOrder.sort(), classes: built };
}

function seasonDriverKey(entry) {
  return normalizeKey(`${entry.driver}|${entry.displayClass || entry.class || entry.cls || ''}`);
}

function renderPointsClassCard(cls, rows, points) {
  return `
    <section class="card results-card" data-class="${escapeHtml(cls)}">
      <div class="card-header">
        <div class="class-title">
          <div class="acr-tag">${escapeHtml(cls)}</div>
          <div class="header-main">Points Standings</div>
          <div class="driver-sub">Net total after ${points.dropCount} lowest point day${points.dropCount === 1 ? '' : 's'} dropped</div>
        </div>
        <div class="class-count">${rows.length} DRIVER${rows.length === 1 ? '' : 'S'}</div>
      </div>
      <div class="card-body">
        <div class="points-table-wrap points-table-compact-wrap">
          <table class="compact-table points-table points-table-rounds-under">
            <thead>
              <tr><th>#</th><th>Driver / Round Scores</th><th>Net</th><th>Gross</th><th>Dropped</th><th>Events</th><th>Wins</th><th>Podiums</th></tr>
            </thead>
            <tbody>
              ${rows.map(row => renderPointsRow(row)).join('') || `<tr><td colspan="8">No drivers found.</td></tr>`}
            </tbody>
          </table>
        </div>
      </div>
    </section>
  `;
}


function getAllWorkbookData() {
  const rounds = state.workbook?.rounds || [];
  const previous = state.selectedRoundIds;
  state.selectedRoundIds = rounds.map(round => round.id);
  const data = getSelectedData();
  state.selectedRoundIds = previous;
  return data || { selectedRounds: rounds, entries: [], classes: {}, classOrder: [], overall: [], pax: [], summary: [] };
}

function buildDriverOptionsFromSummary(summary) {
  return summary.map(row => ({ ...row, label: labelForDriver(row) }));
}

function renderDriverInsights(data) {
  const root = document.getElementById('rankings');
  const seasonData = getAllWorkbookData();
  const driverOptions = buildDriverOptionsFromSummary(seasonData.summary);
  const selectedDriver = findDriverInOptions(state.insightDriverSelection, driverOptions);
  const selectedRival = findDriverInOptions(state.insightRivalSelection, driverOptions);
  const selectedRival2 = findDriverInOptions(state.insightRivalSelection2, driverOptions);

  root.innerHTML = `
    <section class="results-shell">
      ${renderViewDock(data)}
      <section class="card results-card compare-shell insights-shell">
        <div class="card-header">
          <div class="class-title">
            <div class="acr-tag">INSIGHTS</div>
            <div class="header-main">Driver Insights</div>
            <div class="driver-sub">Season analysis by driver and rivals.</div>
          </div>
          <div class="class-count">${seasonData.selectedRounds.length} EVENT${seasonData.selectedRounds.length === 1 ? '' : 'S'}</div>
        </div>
        <div class="card-body">
          <datalist id="insightDriverOptions">${driverOptions.map(driver => `<option value="${escapeHtml(driver.label)}"></option>`).join('')}</datalist>
          <div class="compare-input-grid">
            <label class="compare-input-wrap"><span>Select Driver</span><input class="compare-input" data-insight-driver list="insightDriverOptions" placeholder="Start typing a driver name..." value="${escapeHtml(state.insightDriverSelection || '')}" /></label>
            <label class="compare-input-wrap"><span>Driver 2 / Rival</span><input class="compare-input" data-insight-rival list="insightDriverOptions" placeholder="Optional rival benchmark..." value="${escapeHtml(state.insightRivalSelection || '')}" /></label>
            <label class="compare-input-wrap"><span>Driver 3 / Rival</span><input class="compare-input" data-insight-rival2 list="insightDriverOptions" placeholder="Optional third driver..." value="${escapeHtml(state.insightRivalSelection2 || '')}" /></label>
          </div>
          ${selectedDriver ? renderDriverInsightContent(seasonData, selectedDriver, [selectedRival, selectedRival2].filter(Boolean)) : `<div class="compare-empty">Select a driver to see season snapshot, event benchmarks, run development, and closest battles.</div>`}
        </div>
      </section>
    </section>
  `;
  attachInsightHandlers();
}

function findDriverInOptions(label, options) {
  const wanted = normalizeKey(label);
  if (!wanted) return null;
  return options.find(driver => normalizeKey(driver.label) === wanted) ||
    options.find(driver => normalizeKey(driver.driver) === wanted) ||
    options.find(driver => normalizeKey(driver.label).includes(wanted));
}

function attachInsightHandlers() {
  document.querySelectorAll('[data-insight-driver]').forEach(input => {
    input.addEventListener('input', event => { state.insightDriverSelection = event.target.value; });
    input.addEventListener('change', event => { state.insightDriverSelection = event.target.value; render(); });
    input.addEventListener('keydown', event => { if (event.key === 'Enter') { event.target.blur(); render(); } });
  });
  document.querySelectorAll('[data-insight-rival]').forEach(input => {
    input.addEventListener('input', event => { state.insightRivalSelection = event.target.value; });
    input.addEventListener('change', event => { state.insightRivalSelection = event.target.value; render(); });
    input.addEventListener('keydown', event => { if (event.key === 'Enter') { event.target.blur(); render(); } });
  });
  document.querySelectorAll('[data-insight-rival2]').forEach(input => {
    input.addEventListener('input', event => { state.insightRivalSelection2 = event.target.value; });
    input.addEventListener('change', event => { state.insightRivalSelection2 = event.target.value; render(); });
    input.addEventListener('keydown', event => { if (event.key === 'Enter') { event.target.blur(); render(); } });
  });
}

function renderDriverInsightContent(data, selectedDriver, selectedRivals = []) {
  const selectedKey = driverKey(selectedDriver);
  const rivalKeys = selectedRivals.map(rival => driverKey(rival));
  const rows = data.entries
    .filter(row => driverKey(row) === selectedKey)
    .slice()
    .sort((a, b) => roundSortValue(a.round) - roundSortValue(b.round));
  const rivalRowsList = rivalKeys.map(key => data.entries.filter(row => driverKey(row) === key));
  return `
    <div class="insight-section">
      <div class="section-label">Season Snapshot</div>
      ${assumptionNote([
        'Total points are pulled from the season points calculation for each driver’s parsed event group/class.',
        'Clean run percentage counts parsed run attempts marked clean, excluding cones, DNF, and RRN/rerun entries.',
        'Rival cards use the same calculation so drivers can compare snapshots side by side.'
      ])}
      ${renderSeasonSnapshotComparison(data, [selectedDriver, ...selectedRivals].slice(0, 3))}
    </div>

    <div class="insight-section">
      <div class="section-label">Best / Worst Event Breakdown</div>
      ${assumptionNote([
        'Best raw and PAX ranks are event-overall ranks within that round.',
        'Best class rank and worst event use the driver’s parsed class/group finish.',
        'Most improved event compares first clean raw run to best clean raw run.',
        'Dropped event uses the season points calculation.'
      ])}
      ${renderBestWorstBreakdown(data, [selectedDriver, ...selectedRivals].slice(0, 3))}
    </div>

    ${selectedRivals.length ? renderComparePointsTrend([selectedDriver, ...selectedRivals].slice(0, 3)) : ''}

    <div class="insight-section">
      <div class="section-label">Event-by-Event Benchmark</div>
      ${assumptionNote([
        'Benchmarks use best raw time for the selected event group, for example S3, N-, or S2.',
        'Class means the base or PAX class, for example CS, BS, or FS. Group means the event grouping, for example S3, N-, or BST.',
        'Group winner and median are calculated from the displayed group table for that event, not the whole event overall.',
        'Selected rival margins compare event-best raw times. They are not run-by-run comparisons.'
      ])}
      ${renderEventBenchmarkTable(data, rows, rivalRowsList, selectedRivals)}
    </div>

    <div class="insight-section">
      <div class="section-label">Run Development</div>
      ${assumptionNote([
        'First clean, best clean, improvement, and progression use raw run times only.',
        'Cone runs, DNFs, and RRN/rerun entries are excluded from clean-best progression.',
        'Progression only shows a new time when the driver improves their clean best.'
      ])}
      ${renderRunDevelopmentTable(rows)}
    </div>

    <div class="insight-section">
      <div class="section-label">Closest Battles</div>
      ${assumptionNote([
        'Closest battles use event-best raw time margins inside the selected event group.',
        'This is not PAX-normalized unless a separate PAX margin is shown.',
        'Closest event means the closest event-best result, not the closest individual run during that event.'
      ])}
      ${renderClosestBattles(data, selectedDriver)}
    </div>
  `;
}


function renderBestWorstBreakdown(data, drivers) {
  const cards = drivers.map((driver, index) => renderBestWorstCard(buildBestWorstSummary(data, driver), index)).join('');
  return `<div class="bestworst-grid bestworst-count-${drivers.length}">${cards}</div>`;
}

function buildBestWorstSummary(data, driver) {
  const key = driverKey(driver);
  const rows = data.entries
    .filter(row => driverKey(row) === key)
    .slice()
    .sort((a, b) => roundSortValue(a.round) - roundSortValue(b.round));
  const seasonPoints = findSeasonPointsForDriver(driver);
  const scoredRows = rows.filter(row => Number.isFinite(Number(row.bestRaw)));

  const bestRaw = scoredRows
    .map(row => ({ row, rank: eventRawRank(data, row) }))
    .filter(item => Number.isFinite(Number(item.rank)))
    .sort((a, b) => Number(a.rank) - Number(b.rank))[0];

  const bestPax = scoredRows
    .map(row => ({ row, rank: eventPaxRank(data, row) }))
    .filter(item => Number.isFinite(Number(item.rank)))
    .sort((a, b) => Number(a.rank) - Number(b.rank))[0];

  const bestClass = scoredRows
    .filter(row => Number.isFinite(Number(row.position)))
    .sort((a, b) => Number(a.position) - Number(b.position))[0];

  const worstEvent = scoredRows
    .filter(row => Number.isFinite(Number(row.position)))
    .sort((a, b) => Number(b.position) - Number(a.position) || Number(b.bestRaw || 0) - Number(a.bestRaw || 0))[0];

  const mostImproved = rows
    .map(row => {
      const clean = cleanRunTimes(row);
      const firstClean = clean.length ? clean[0] : null;
      const bestClean = clean.length ? Math.min(...clean) : null;
      const improvement = Number.isFinite(firstClean) && Number.isFinite(bestClean) ? firstClean - bestClean : null;
      return { row, improvement };
    })
    .filter(item => Number.isFinite(Number(item.improvement)))
    .sort((a, b) => Number(b.improvement) - Number(a.improvement))[0];

  const droppedEvents = seasonPoints?.eventScores
    ?.map((score, index) => ({ ...score, index }))
    ?.filter(score => seasonPoints.droppedIndexes?.has(score.index)) || [];

  return {
    driver,
    bestRaw,
    bestPax,
    bestClass,
    worstEvent,
    mostImproved,
    droppedEvents
  };
}

function eventRawRank(data, row) {
  const rows = data.entries
    .filter(entry => entry.round?.id === row.round?.id)
    .filter(entry => Number.isFinite(Number(entry.bestRaw)))
    .slice()
    .sort((a, b) => Number(a.bestRaw) - Number(b.bestRaw));
  const index = rows.findIndex(entry => driverKey(entry) === driverKey(row));
  return index >= 0 ? index + 1 : null;
}

function eventPaxRank(data, row) {
  const rows = data.entries
    .filter(entry => entry.round?.id === row.round?.id)
    .filter(entry => Number.isFinite(Number(entry.bestPax)))
    .slice()
    .sort((a, b) => Number(a.bestPax) - Number(b.bestPax));
  const index = rows.findIndex(entry => driverKey(entry) === driverKey(row));
  return index >= 0 ? index + 1 : null;
}

function eventLabelFromRow(row) {
  return row?.round?.label || row?.round?.id || '—';
}

function renderBestWorstCard(summary, index) {
  const dropped = summary.droppedEvents?.length
    ? summary.droppedEvents.map(score => `${score.round?.label || score.round?.id || 'Event'} · ${score.missed ? 'Miss' : `${score.points} pts / P${score.place}`}`).join(', ')
    : '—';
  return `
    <article class="bestworst-card bestworst-card-${index + 1}">
      <div class="snapshot-role">${escapeHtml(index === 0 ? 'Selected Driver' : `Rival ${index}`)}</div>
      <div class="driver-name">${escapeHtml(summary.driver.driver)} ${escapeHtml(summary.driver.number || '')}</div>
      <div class="driver-sub">${escapeHtml(driverContextLabel(summary.driver, false))}</div>
      <div class="bestworst-metric-grid">
        ${bestWorstMetric('Best Raw Rank', summary.bestRaw ? `P${summary.bestRaw.rank}` : '—', summary.bestRaw ? eventLabelFromRow(summary.bestRaw.row) : '')}
        ${bestWorstMetric('Best PAX Rank', summary.bestPax ? `P${summary.bestPax.rank}` : '—', summary.bestPax ? eventLabelFromRow(summary.bestPax.row) : '')}
        ${bestWorstMetric('Best Class Rank', summary.bestClass ? `P${summary.bestClass.position}` : '—', summary.bestClass ? eventLabelFromRow(summary.bestClass) : '')}
        ${bestWorstMetric('Worst Event', summary.worstEvent ? `P${summary.worstEvent.position}` : '—', summary.worstEvent ? eventLabelFromRow(summary.worstEvent) : '')}
        ${bestWorstMetric('Most Improved', summary.mostImproved ? `${Number(summary.mostImproved.improvement).toFixed(3)}s` : '—', summary.mostImproved ? eventLabelFromRow(summary.mostImproved.row) : '')}
        ${bestWorstMetric('Dropped Event', dropped, '')}
      </div>
    </article>
  `;
}

function bestWorstMetric(label, value, sub) {
  return `<div class="bestworst-metric"><span>${escapeHtml(label)}</span><strong>${escapeHtml(value ?? '—')}</strong>${sub ? `<em>${escapeHtml(sub)}</em>` : ''}</div>`;
}


function buildInsightSnapshot(data, driver, roleLabel) {
  const key = driverKey(driver);
  const rows = data.entries
    .filter(row => driverKey(row) === key)
    .slice()
    .sort((a, b) => roundSortValue(a.round) - roundSortValue(b.round));
  const seasonPoints = findSeasonPointsForDriver(driver);
  const runStats = collectRunStats(rows);
  const positions = rows.map(row => Number(row.position)).filter(Number.isFinite);
  const bestFinish = positions.length ? Math.min(...positions) : null;
  const avgFinish = positions.length ? (positions.reduce((a, b) => a + b, 0) / positions.length).toFixed(1) : '—';
  const cleanPct = runStats.totalRuns ? `${Math.round((runStats.cleanRuns / runStats.totalRuns) * 100)}%` : '—';

  return {
    driver,
    roleLabel,
    rows,
    eventsRun: rows.length,
    totalPoints: seasonPoints ? seasonPoints.netPoints : '—',
    grossPoints: seasonPoints ? seasonPoints.grossPoints : '—',
    bestFinish: bestFinish ? `P${bestFinish}` : '—',
    avgFinish,
    cleanPct,
    coneDnf: `${runStats.coneRuns} / ${runStats.dnfRuns}`
  };
}

function renderSeasonSnapshotComparison(data, drivers) {
  const snapshots = drivers.map((driver, index) => buildInsightSnapshot(data, driver, index === 0 ? 'Selected Driver' : `Rival ${index}`));
  return `
    <div class="snapshot-card-grid snapshot-count-${snapshots.length}">
      ${snapshots.map((snapshot, index) => renderSeasonSnapshotCard(snapshot, index)).join('')}
    </div>
  `;
}

function renderSeasonSnapshotCard(snapshot, index) {
  return `
    <article class="snapshot-card snapshot-card-${index + 1}">
      <div class="snapshot-head">
        <div>
          <div class="snapshot-role">${escapeHtml(snapshot.roleLabel)}</div>
          <div class="driver-name">${escapeHtml(snapshot.driver.driver)} ${escapeHtml(snapshot.driver.number || '')}</div>
          <div class="driver-sub">${escapeHtml(driverContextLabel(snapshot.driver, false))}</div>
        </div>
        <div class="snapshot-points">
          <strong>${escapeHtml(snapshot.totalPoints)}</strong>
          <span>net pts</span>
        </div>
      </div>
      <div class="snapshot-metric-grid">
        ${snapshotMetric('Events', snapshot.eventsRun)}
        ${snapshotMetric('Gross', snapshot.grossPoints)}
        ${snapshotMetric('Best', snapshot.bestFinish)}
        ${snapshotMetric('Avg Finish', snapshot.avgFinish)}
        ${snapshotMetric('Clean', snapshot.cleanPct)}
        ${snapshotMetric('Cones / DNF', snapshot.coneDnf)}
      </div>
    </article>
  `;
}

function snapshotMetric(label, value) {
  return `<div class="snapshot-metric"><span>${escapeHtml(label)}</span><strong>${escapeHtml(value ?? '—')}</strong></div>`;
}

function collectRunStats(rows) {
  const runs = rows.flatMap(row => row.runObjects || []);
  return {
    totalRuns: runs.length,
    cleanRuns: runs.filter(run => run.isClean).length,
    coneRuns: runs.filter(run => run.hasCone).length,
    dnfRuns: runs.filter(run => run.isDnf).length,
    rerunRuns: runs.filter(run => run.isRerun).length
  };
}

function findSeasonPointsForDriver(driver) {
  const points = buildSeasonPoints();
  const cls = driver.displayClass || driver.class || driver.cls || 'UNKNOWN';
  const wanted = seasonDriverKey(driver);
  return (points.classes[cls] || []).find(row => seasonDriverKey(row) === wanted) || null;
}

function classRowsForEvent(data, row) {
  const cls = row.displayClass || row.class || row.cls || 'UNKNOWN';
  return data.entries
    .filter(entry => entry.round?.id === row.round?.id && (entry.displayClass || entry.class || entry.cls || 'UNKNOWN') === cls)
    .filter(entry => Number.isFinite(Number(entry.bestRaw)))
    .slice()
    .sort((a, b) => Number(a.bestRaw) - Number(b.bestRaw));
}

function median(values) {
  const nums = values.map(Number).filter(Number.isFinite).sort((a, b) => a - b);
  if (!nums.length) return null;
  const mid = Math.floor(nums.length / 2);
  return nums.length % 2 ? nums[mid] : (nums[mid - 1] + nums[mid]) / 2;
}

function renderEventBenchmarkTable(data, rows, rivalRowsList = [], selectedRivals = []) {
  const rivalMaps = rivalRowsList.map(rivalRows => new Map(rivalRows.map(row => [row.round?.id, row])));
  const rivalHeaders = selectedRivals.map((rival, index) => `<th>Vs ${escapeHtml(rival.driver || `Rival ${index + 1}`)}</th>`).join('');
  const body = rows.map(row => {
    const classRows = classRowsForEvent(data, row);
    const winner = classRows[0];
    const med = median(classRows.map(entry => entry.bestRaw));
    const rivalCells = rivalMaps.map((rivalByRound, index) => {
      const rival = rivalByRound.get(row.round?.id);
      return `<td>${escapeHtml(rival ? formatTime(rival.bestRaw) : '—')}<span class="table-sub">${escapeHtml(rival ? formatBattleMargin(Number(rival.bestRaw) - Number(row.bestRaw)) : 'NO RESULT')}</span></td>`;
    }).join('');
    return `
      <tr>
        <td>${escapeHtml(row.round?.label || '—')}</td>
        <td>${escapeHtml(baseClassLabel(row) || '—')}<span class="table-sub">CLASS</span></td>
        <td>${escapeHtml(groupLabel(row))}<span class="table-sub">GROUP</span></td>
        <td><strong>${escapeHtml(formatTime(row.bestRaw))}</strong><span class="table-sub">P${escapeHtml(row.position ?? '—')}</span></td>
        <td>${escapeHtml(winner ? winner.driver : '—')}<span class="table-sub">${escapeHtml(winner ? formatGap(Number(row.bestRaw) - Number(winner.bestRaw)) : '—')}</span></td>
        <td>${escapeHtml(formatTime(med))}<span class="table-sub">${escapeHtml(Number.isFinite(Number(row.bestRaw)) && Number.isFinite(Number(med)) ? formatSignedGap(Number(row.bestRaw) - Number(med)) : '—')}</span></td>
        ${rivalCells || `<td>—<span class="table-sub">NO RIVAL</span></td>`}
      </tr>
    `;
  }).join('');

  const colspan = 6 + Math.max(1, selectedRivals.length);
  return `
    <div class="points-table-wrap">
      <table class="compact-table insight-table">
        <thead><tr><th>Event</th><th>Class</th><th>Group</th><th>Driver Raw</th><th>Vs Group Winner</th><th>Vs Group Median</th>${rivalHeaders || '<th>Vs Selected Rival</th>'}</tr></thead>
        <tbody>${body || `<tr><td colspan="${colspan}">No benchmark rows found.</td></tr>`}</tbody>
      </table>
    </div>
  `;
}

function formatSignedGap(value) {
  const num = Number(value);
  if (!Number.isFinite(num)) return '—';
  if (Math.abs(num) < 0.0005) return 'EVEN';
  return `${num > 0 ? '+' : ''}${num.toFixed(3)}`;
}

function cleanRunTimes(row) {
  return (row.runObjects || [])
    .filter(run => run.isClean && Number.isFinite(Number(run.rawTime)))
    .map(run => Number(run.rawTime));
}

function cleanBestProgression(row) {
  let best = Infinity;
  const progression = [];
  (row.runObjects || []).forEach(run => {
    const time = Number(run.rawTime);
    if (!run.isClean || !Number.isFinite(time)) return;
    if (time < best - 0.0005) {
      best = time;
      progression.push(time);
    }
  });
  return progression;
}

function renderRunDevelopmentTable(rows) {
  const body = rows.map(row => {
    const clean = cleanRunTimes(row);
    const firstClean = clean.length ? clean[0] : null;
    const bestClean = clean.length ? Math.min(...clean) : null;
    const bestProgression = cleanBestProgression(row).map(time => formatTime(time)).join(' → ');
    const improvement = Number.isFinite(firstClean) && Number.isFinite(bestClean) ? firstClean - bestClean : null;
    return `
      <tr>
        <td>${escapeHtml(row.round?.label || '—')}</td>
        <td>${escapeHtml(formatTime(firstClean))}</td>
        <td>${escapeHtml(formatTime(bestClean))}</td>
        <td>${escapeHtml(Number.isFinite(Number(improvement)) ? improvement.toFixed(3) : '—')}</td>
        <td>${escapeHtml(bestProgression || '—')}</td>
        <td>${(row.runObjects || []).map(run => `<span class="run-pill ${runClass(run)}">${escapeHtml(run.label)}</span>`).join('')}</td>
      </tr>
    `;
  }).join('');

  return `
    <div class="points-table-wrap">
      <table class="compact-table insight-table">
        <thead><tr><th>Event</th><th>First Clean</th><th>Best Clean</th><th>Improvement</th><th>Clean Best Progression</th><th>Runs</th></tr></thead>
        <tbody>${body || `<tr><td colspan="6">No run development rows found.</td></tr>`}</tbody>
      </table>
    </div>
  `;
}

function runningBest(times) {
  let best = Infinity;
  return times.map(time => {
    if (time < best) best = time;
    return best;
  });
}

function formatBattleMargin(margin) {
  if (!Number.isFinite(Number(margin))) return '—';
  const value = Math.abs(Number(margin)).toFixed(3);
  if (Math.abs(Number(margin)) < 0.0005) return 'Even';
  return Number(margin) > 0 ? `Driver by ${value}` : `Rival by ${value}`;
}

function renderClosestBattles(data, selectedDriver) {
  const selectedKey = driverKey(selectedDriver);
  const className = selectedDriver.displayClass || selectedDriver.class || selectedDriver.cls || 'UNKNOWN';
  const selectedRows = data.entries.filter(row => driverKey(row) === selectedKey);
  const battleMap = new Map();

  selectedRows.forEach(row => {
    if (!Number.isFinite(Number(row.bestRaw))) return;
    data.entries
      .filter(other => other.round?.id === row.round?.id)
      .filter(other => driverKey(other) !== selectedKey)
      .filter(other => (other.displayClass || other.class || other.cls || 'UNKNOWN') === className)
      .filter(other => Number.isFinite(Number(other.bestRaw)))
      .forEach(other => {
        const margin = Number(other.bestRaw) - Number(row.bestRaw); // positive means selected driver was quicker
        const paxMargin = Number.isFinite(Number(other.bestPax)) && Number.isFinite(Number(row.bestPax)) ? Number(other.bestPax) - Number(row.bestPax) : null;
        const absGap = Math.abs(margin);
        const key = driverKey(other);
        const item = battleMap.get(key) || {
          driver: other.driver,
          number: other.number,
          car: other.car,
          displayClass: className,
          baseClass: baseClassLabel(other),
          matchups: 0,
          closeEvents: 0,
          margins: [],
          paxMargins: [],
          closest: Infinity,
          closestMargin: null,
          closestEvent: ''
        };

        item.matchups += 1;
        item.margins.push(margin);
        if (Number.isFinite(Number(paxMargin))) item.paxMargins.push(paxMargin);

        if (absGap <= 1) {
          item.closeEvents += 1;
        }

        if (absGap < item.closest) {
          item.closest = absGap;
          item.closestMargin = margin;
          item.closestEvent = row.round?.label || row.round?.name || row.round?.id || 'Event';
        }

        battleMap.set(key, item);
      });
  });

  const battles = Array.from(battleMap.values())
    .filter(item => item.closeEvents > 0)
    .map(item => ({
      ...item,
      avgMargin: item.margins.reduce((sum, margin) => sum + margin, 0) / item.margins.length,
      avgPaxMargin: item.paxMargins.length ? item.paxMargins.reduce((sum, margin) => sum + margin, 0) / item.paxMargins.length : null
    }))
    .sort((a, b) => b.closeEvents - a.closeEvents || Math.abs(a.avgMargin) - Math.abs(b.avgMargin) || a.closest - b.closest)
    .slice(0, 8);

  if (!battles.length) {
    return `<div class="compare-empty">No same-group drivers found within 1.000 second across the season.</div>`;
  }

  const cards = battles.map((item, index) => `
    <article class="battle-card">
      <div class="battle-rank rank ${rankClass(index + 1)}">${index + 1}</div>
      <div class="battle-main">
        <div class="battle-driver-row">
          <div>
            <div class="driver-name">${escapeHtml(item.driver)} ${escapeHtml(item.number || '')}</div>
            <div class="driver-sub">${escapeHtml([item.displayClass, item.baseClass, item.car].filter(Boolean).join(' · '))}</div>
          </div>
          <div class="battle-close">
            <strong>${escapeHtml(item.closest.toFixed(3))}</strong>
            <span>closest raw</span>
          </div>
        </div>
        <div class="battle-meta-grid">
          <div class="battle-chip"><span>Close events</span><strong>${escapeHtml(item.closeEvents)} of ${escapeHtml(item.matchups)}</strong></div>
          <div class="battle-chip"><span>Closest event</span><strong>${escapeHtml(item.closestEvent)} · ${escapeHtml(formatBattleMargin(item.closestMargin))}</strong></div>
          <div class="battle-chip"><span>Avg raw margin</span><strong>${escapeHtml(formatBattleMargin(item.avgMargin))}</strong></div>
          <div class="battle-chip"><span>Avg PAX margin</span><strong>${escapeHtml(formatBattleMargin(item.avgPaxMargin))}</strong></div>
        </div>
      </div>
    </article>
  `).join('');

  return `
    <div class="battle-legend"><strong>RAW</strong> same-group event-best time. <strong>PAX</strong> reference only when both drivers have PAX values. <strong>Close</strong> means within 1.000 second.</div>
    <div class="battle-card-list">${cards}</div>
  `;
}

function renderPointsRow(row) {
  const roundScores = row.eventScores.map((score, index) => {
    const dropped = row.droppedIndexes.has(index);
    const label = score.missed ? 'MISS' : `P${score.place}`;
    const roundLabel = score.round?.label || score.round?.id || `R${index + 1}`;
    const cls = [score.missed ? 'points-miss' : '', dropped ? 'points-drop' : ''].filter(Boolean).join(' ');
    return `
      <span class="points-score-chip ${cls}" title="${escapeHtml(roundLabel)} · ${escapeHtml(score.missed ? 'Missed' : `${score.points} points · ${label}`)}${dropped ? ' · Dropped' : ''}">
        <span class="points-round-label">${escapeHtml(shortRoundLabel(roundLabel, index))}</span>
        <strong>${escapeHtml(score.points)}</strong>
        <span>${escapeHtml(label)}</span>
      </span>
    `;
  }).join('');

  return `
    <tr class="points-main-row">
      <td><div class="rank ${rankClass(row.rank)}">${escapeHtml(row.rank)}</div></td>
      <td class="points-driver-cell">
        <div class="driver-name">${escapeHtml(row.driver)} ${escapeHtml(row.number || '')}</div>
        <div class="driver-sub">${escapeHtml([row.displayClass, row.car].filter(Boolean).join(' · '))}</div>
      </td>
      <td><strong>${escapeHtml(row.netPoints)}</strong></td>
      <td>${escapeHtml(row.grossPoints)}</td>
      <td>${escapeHtml(row.droppedPoints)}</td>
      <td>${escapeHtml(row.eventsRun)}</td>
      <td>${escapeHtml(row.wins)}</td>
      <td>${escapeHtml(row.podiums)}</td>
    </tr>
    <tr class="points-rounds-subrow">
      <td></td>
      <td class="points-rounds-fullcell" colspan="7">
        <div class="points-round-strip">${roundScores}</div>
      </td>
    </tr>
  `;
}

function shortRoundLabel(label, index) {
  const match = String(label || '').match(/(\d+)/);
  return match ? `R${match[1]}` : `R${index + 1}`;
}

function renderResultRow(row, timeLabel, timeValue) {
  return `
    <div class="result-row">
      <div class="rank ${rankClass(row.rank || row.position)}">${escapeHtml(row.rank || row.position || '—')}</div>
      <div>
        <div class="driver-name">${escapeHtml(row.driver)} ${escapeHtml(row.number || '')}</div>
        <div class="driver-sub">${escapeHtml([row.round?.label, row.displayClass || row.class, row.car].filter(Boolean).join(' · '))}</div>
        ${row.runs?.length ? `<div class="run-strip">${row.runObjects.map(run => `<span class="run-pill ${runClass(run)}">${escapeHtml(run.label)}</span>`).join('')}</div>` : ''}
      </div>
      <div class="time-cell">
        <span class="time-val">${escapeHtml(formatTime(timeValue))}</span>
        <span class="time-label">${escapeHtml(timeLabel)}</span>
      </div>
    </div>
  `;
}

function runClass(run) {
  if (run.isDnf) return 'dnf';
  if (run.isRerun) return 'rerun';
  if (run.hasCone) return 'cone';
  return '';
}

function rankClass(rank) {
  const value = Number(rank);
  return value <= 3 ? `rank-${value}` : '';
}

function emptyRow(message) {
  return `<div class="result-row"><div class="rank">—</div><div><div class="driver-name">${escapeHtml(message)}</div><div class="driver-sub">Choose an Excel workbook and select rounds to render data.</div></div><div class="time-cell"><span class="time-val">—</span><span class="time-label">NO DATA</span></div></div>`;
}

function renderEmptyCard(tag, message) {
  return `<section class="results-shell"><section class="card results-card"><div class="card-header"><div class="class-title"><div class="acr-tag">${escapeHtml(tag)}</div><div class="header-main">${escapeHtml(message)}</div></div></div><div class="card-body">${emptyRow(message)}</div></section></section>`;
}


function sortedSelectedRounds(data) {
  return [...(data.selectedRounds || [])].sort((a, b) => roundSortValue(b) - roundSortValue(a));
}

function entriesForRound(data, round) {
  return (data.entries || []).filter(row => row.round?.id === round.id);
}

function rankOverallRows(entries) {
  return entries
    .filter(row => Number.isFinite(Number(row.bestRaw)))
    .slice()
    .sort((a, b) => Number(a.bestRaw) - Number(b.bestRaw))
    .map((row, index) => ({ ...row, rank: index + 1, time: row.bestRaw, overallRank: index + 1 }));
}

function rankPaxRows(entries) {
  return entries
    .filter(row => Number.isFinite(Number(row.bestPax)))
    .slice()
    .sort((a, b) => Number(a.bestPax) - Number(b.bestPax))
    .map((row, index) => ({ ...row, rank: index + 1, time: row.bestPax, paxRank: index + 1 }));
}

function buildSummaryRows(summaryRows) {
  return summaryRows.slice(0, 80).map(row => `
    <tr>
      <td><div class="rank ${rankClass(row.rank)}">${escapeHtml(row.rank)}</div></td>
      <td>
        <div class="driver-name">${escapeHtml(row.driver)} ${escapeHtml(row.number || '')}</div>
        <div class="driver-sub">${escapeHtml([row.displayClass || row.class, row.car].filter(Boolean).join(' · '))}</div>
        <div>${[...new Set(row.roundLabels)].map(label => `<span class="round-chip">${escapeHtml(label)}</span>`).join('')}</div>
      </td>
      <td>${escapeHtml(row.roundsCount)}</td>
      <td>${escapeHtml(formatTime(row.bestRaw))}</td>
      <td>${escapeHtml(formatTime(row.bestPax))}</td>
      <td>${escapeHtml(formatTime(row.avgPax))}</td>
      <td>${escapeHtml(row.bestClassPosition ?? '—')}</td>
      <td>${escapeHtml(row.classWins)}</td>
    </tr>
  `).join('');
}

function renderSummaryMetrics(entries, summaryRows, selectedRoundCount, classCount) {
  const cleanRuns = entries.flatMap(row => row.runObjects || []).filter(run => run.isClean).length;
  const dnfRuns = entries.flatMap(row => row.runObjects || []).filter(run => run.isDnf).length;
  const rerunRuns = entries.flatMap(row => row.runObjects || []).filter(run => run.isRerun).length;
  return `
    <div class="summary-grid">
      ${metricCard('ROUNDS', selectedRoundCount)}
      ${metricCard('CLASSES', classCount)}
      ${metricCard('ENTRIES', entries.length)}
      ${metricCard('CLEAN RUNS', cleanRuns)}
      ${metricCard('DNF', dnfRuns)}
      ${metricCard('RRN / RERUN', rerunRuns)}
    </div>
  `;
}

function renderSummaryTable(summaryRows) {
  return `
    <table class="compact-table">
      <thead><tr><th>#</th><th>Driver</th><th>Rounds</th><th>Best Raw</th><th>Best PAX</th><th>Avg PAX</th><th>Best Class</th><th>Wins</th></tr></thead>
      <tbody>${buildSummaryRows(summaryRows) || `<tr><td colspan="8">No summary rows found.</td></tr>`}</tbody>
    </table>
  `;
}

function renderRoundBreakoutHeader(round, countLabel) {
  return `
    <div class="round-group-header">
      <div>
        <div class="round-kicker">EVENT BREAKOUT</div>
        <div class="round-group-title">${escapeHtml(round.label || round.name || round.id)}</div>
        <div class="driver-sub">${escapeHtml([round.date, round.location].filter(Boolean).join(' · '))}</div>
      </div>
      <div class="class-count">${escapeHtml(countLabel)}</div>
    </div>
  `;
}

function renderCompare(data) {
  const root = document.getElementById('rankings');
  const selected = state.compareSelections.map(label => findSelectedDriver(label)).filter(Boolean);
  const multiRound = (data.selectedRounds || []).length > 1;

  root.innerHTML = `
    <section class="results-shell">
      ${renderViewDock(data)}
      <section class="card results-card compare-shell">
        <div class="card-header">
          <div class="class-title"><div class="acr-tag">COMPARE</div><div class="header-main">Driver Comparison</div><div class="driver-sub">${escapeHtml(selectedRoundLabel(data))}</div></div>
          <div class="class-count">${selected.length} SELECTED</div>
        </div>
        <div class="card-body">
          <datalist id="driverOptions">${state.driverIndex.map(driver => `<option value="${escapeHtml(driver.label)}"></option>`).join('')}</datalist>
          <div class="compare-input-grid">
            ${[0, 1, 2].map(index => `
              <label class="compare-input-wrap"><span>Driver ${index + 1}</span><input class="compare-input" data-compare-index="${index}" list="driverOptions" placeholder="Start typing a driver name..." value="${escapeHtml(state.compareSelections[index] || '')}" /></label>
            `).join('')}
          </div>
          ${assumptionNote([
            'Class shows the base or PAX class, for example CS, BS, or FS. Group shows the event grouping, for example S3, S2, N-, or BST.',
            'Raw gaps compare best raw times. PAX gaps compare best PAX times and are better for cross-class interpretation.',
            'When exactly two drivers are selected, the trend chart uses per-event season points across the season, before drops.'
          ])}
          ${selected.length === 2 ? renderComparePointsTrend(selected) : ''}
          ${selected.length ? (multiRound ? renderCompareByRound(data, selected) : renderCompareSingle(selected)) : `<div class="compare-empty">Select two or three drivers to compare selected-round performance.</div>`}
        </div>
      </section>
    </section>
  `;
  attachCompareHandlers();
}

function renderCompareSingle(selected) {
  return `${renderCompareCards(selected)}${selected.length >= 2 ? renderGapAnalysis(selected) : ''}`;
}

function renderCompareByRound(data, selectedDrivers) {
  const selectedKeys = new Set(selectedDrivers.map(driver => driverKey(driver)));
  const sections = sortedSelectedRounds(data).map(round => {
    const rows = entriesForRound(data, round)
      .filter(row => selectedKeys.has(driverKey(row)))
      .sort((a, b) => Number(a.bestPax ?? 9999) - Number(b.bestPax ?? 9999));

    return `
      <section class="round-class-group" data-round="${escapeHtml(round.id)}">
        ${renderRoundBreakoutHeader(round, `${rows.length} MATCH${rows.length === 1 ? '' : 'ES'}`)}
        ${rows.length ? renderCompareEntryCards(rows) : `<div class="compare-empty">None of the selected drivers have a parsed result in this event.</div>`}
        ${rows.length >= 2 ? renderEntryGapAnalysis(rows) : ''}
      </section>
    `;
  }).join('');

  return sections || `<div class="compare-empty">No matching event rows found for the selected drivers.</div>`;
}

function renderCompareEntryCards(rows) {
  return `<div class="compare-card-grid">${rows.map(row => `
    <article class="compare-card">
      <div class="driver-name">${escapeHtml(row.driver)}</div>
      <div class="driver-sub">${escapeHtml(driverContextLabel(row))}</div>
      <div class="compare-stat-grid">
        ${compareStat('Class Pos', row.position ?? '—')}
        ${compareStat('Best Raw', formatTime(row.bestRaw))}
        ${compareStat('Best PAX', formatTime(row.bestPax))}
        ${compareStat('DNF Runs', (row.runObjects || []).filter(run => run.isDnf).length)}
        ${compareStat('RRN / Rerun', (row.runObjects || []).filter(run => run.isRerun).length)}
        ${compareStat('Clean Runs', (row.runObjects || []).filter(run => run.isClean).length)}
      </div>
      ${row.runObjects?.length ? `<div class="run-strip">${row.runObjects.map(run => `<span class="run-pill ${runClass(run)}">${escapeHtml(run.label)}</span>`).join('')}</div>` : ''}
    </article>
  `).join('')}</div>`;
}

function renderEntryGapAnalysis(rows) {
  const rawLeader = Math.min(...rows.map(d => Number(d.bestRaw)).filter(Number.isFinite));
  const paxLeader = Math.min(...rows.map(d => Number(d.bestPax)).filter(Number.isFinite));
  return `
    <div class="gap-table-wrap">
      <table class="gap-table">
        <thead><tr><th>Driver</th><th>Class</th><th>Group</th><th>Best Raw Gap</th><th>Best PAX Gap</th><th>Best Raw</th><th>Best PAX</th></tr></thead>
        <tbody>
          ${rows.map(row => `
            <tr><td>${escapeHtml(row.driver)}</td><td>${escapeHtml(baseClassLabel(row) || '—')}</td><td>${escapeHtml(groupLabel(row))}</td><td>${escapeHtml(formatGap(Number(row.bestRaw) - rawLeader))}</td><td>${escapeHtml(formatGap(Number(row.bestPax) - paxLeader))}</td><td>${escapeHtml(formatTime(row.bestRaw))}</td><td>${escapeHtml(formatTime(row.bestPax))}</td></tr>
          `).join('')}
        </tbody>
      </table>
    </div>
  `;
}

function findSelectedDriver(label) {
  const wanted = normalizeKey(label);
  if (!wanted) return null;
  return state.driverIndex.find(driver => normalizeKey(driver.label) === wanted) ||
    state.driverIndex.find(driver => normalizeKey(driver.driver) === wanted) ||
    state.driverIndex.find(driver => normalizeKey(driver.label).includes(wanted));
}

function attachCompareHandlers() {
  document.querySelectorAll('.compare-input').forEach(input => {
    input.addEventListener('input', event => {
      state.compareSelections[Number(event.target.dataset.compareIndex)] = event.target.value;
    });
    input.addEventListener('change', event => {
      state.compareSelections[Number(event.target.dataset.compareIndex)] = event.target.value;
      render();
    });
    input.addEventListener('keydown', event => {
      if (event.key === 'Enter') {
        event.target.blur();
        render();
      }
    });
  });
}

function renderCompareCards(drivers) {
  return `<div class="compare-card-grid">${drivers.map(driver => `
    <article class="compare-card">
      <div class="driver-name">${escapeHtml(driver.driver)}</div>
      <div class="driver-sub">${escapeHtml(driverContextLabel(driver))}</div>
      <div class="compare-stat-grid">
        ${compareStat('Rounds', driver.roundsCount)}
        ${compareStat('Best Raw', formatTime(driver.bestRaw))}
        ${compareStat('Best PAX', formatTime(driver.bestPax))}
        ${compareStat('Avg PAX', formatTime(driver.avgPax))}
        ${compareStat('Class Wins', driver.classWins)}
        ${compareStat('DNF Runs', driver.dnfRuns)}
      </div>
      <div>${[...new Set(driver.roundLabels)].map(label => `<span class="round-chip">${escapeHtml(label)}</span>`).join('')}</div>
    </article>
  `).join('')}</div>`;
}

function compareStat(label, value) {
  return `<div class="compare-stat"><span>${escapeHtml(label)}</span><strong>${escapeHtml(value ?? '—')}</strong></div>`;
}

function renderGapAnalysis(drivers) {
  const rawLeader = Math.min(...drivers.map(d => Number(d.bestRaw)).filter(Number.isFinite));
  const paxLeader = Math.min(...drivers.map(d => Number(d.bestPax)).filter(Number.isFinite));
  const avgPaxLeader = Math.min(...drivers.map(d => Number(d.avgPax)).filter(Number.isFinite));
  return `
    <div class="gap-table-wrap">
      <table class="gap-table">
        <thead><tr><th>Driver</th><th>Best Raw Gap</th><th>Best PAX Gap</th><th>Avg PAX Gap</th><th>Selected Rounds</th></tr></thead>
        <tbody>
          ${drivers.map(driver => `
            <tr><td>${escapeHtml(driver.driver)}</td><td>${escapeHtml(formatGap(Number(driver.bestRaw) - rawLeader))}</td><td>${escapeHtml(formatGap(Number(driver.bestPax) - paxLeader))}</td><td>${escapeHtml(formatGap(Number(driver.avgPax) - avgPaxLeader))}</td><td>${escapeHtml(driver.roundsCount)}</td></tr>
          `).join('')}
        </tbody>
      </table>
    </div>
  `;
}


function renderComparePointsTrend(selectedDrivers) {
  if (!selectedDrivers || selectedDrivers.length < 2) return '';
  const points = buildSeasonPoints();
  if (!points.rounds.length) return '';

  const series = selectedDrivers.map(driver => {
    const season = findSeasonPointsForDriver(driver);
    const scores = points.rounds.map((round, index) => {
      const score = season?.eventScores?.[index] || { round, points: 0, place: null, missed: true };
      return {
        round,
        points: Number(score.points || 0),
        place: Number.isFinite(Number(score.place)) ? Number(score.place) : null,
        missed: !!score.missed
      };
    });
    return {
      driver: driver.driver,
      context: driverContextLabel(driver),
      grossPoints: season?.grossPoints ?? 0,
      netPoints: season?.netPoints ?? 0,
      scores
    };
  });

  return `
    <section class="insight-section compare-trend-section">
      <div class="section-label">Season Points Trend</div>
      ${assumptionNote([
        'Y-axis is points earned per event before drop-day removal.',
        'Missed events are plotted as 0 points.',
        'Higher overall line area generally means more gross season points.',
        'Supports up to three selected drivers: the main driver plus two rivals.'
      ])}
      <div class="compare-trend-summary">
        ${series.map((item, index) => `
          <div class="compare-trend-driver compare-trend-driver-${index + 1}">
            <div class="driver-name">${escapeHtml(item.driver)}</div>
            <div class="driver-sub">${escapeHtml(item.context)}</div>
            <div class="compare-trend-meta">
              <span><strong>${escapeHtml(String(item.grossPoints))}</strong> Gross</span>
              <span><strong>${escapeHtml(String(item.netPoints))}</strong> Net</span>
            </div>
          </div>
        `).join('')}
      </div>
      ${renderPointsTrendSvg(points.rounds, series)}
    </section>
  `;
}

function renderPointsTrendSvg(rounds, series) {
  const width = 860;
  const height = 260;
  const margin = { top: 22, right: 22, bottom: 48, left: 46 };
  const innerWidth = width - margin.left - margin.right;
  const innerHeight = height - margin.top - margin.bottom;
  const maxY = Math.max(200, ...series.flatMap(item => item.scores.map(score => Number(score.points || 0))));
  const safeRounds = Math.max(rounds.length - 1, 1);
  const xFor = index => margin.left + (innerWidth * (rounds.length === 1 ? 0.5 : index / safeRounds));
  const yFor = value => margin.top + innerHeight - ((Number(value || 0) / maxY) * innerHeight);
  const yTicks = [0, 50, 100, 150, 200].filter(value => value <= maxY || value === 0);
  if (!yTicks.includes(maxY)) yTicks.push(maxY);

  const colors = ['#dfff00', '#e2e2e2', '#ff9f32'];
  const grid = yTicks.map(value => {
    const y = yFor(value);
    return `<g><line x1="${margin.left}" y1="${y.toFixed(1)}" x2="${(width - margin.right).toFixed(1)}" y2="${y.toFixed(1)}" stroke="rgba(255,255,255,0.10)" stroke-width="1" /><text x="${(margin.left - 8).toFixed(1)}" y="${(y + 4).toFixed(1)}" text-anchor="end" fill="rgba(245,245,245,0.55)" font-size="10">${escapeHtml(String(value))}</text></g>`;
  }).join('');

  const xLabels = rounds.map((round, index) => {
    const x = xFor(index);
    return `<text x="${x.toFixed(1)}" y="${(height - 18).toFixed(1)}" text-anchor="middle" fill="rgba(245,245,245,0.6)" font-size="10">${escapeHtml(round.label || round.id || `R${index + 1}`)}</text>`;
  }).join('');

  const lines = series.map((item, seriesIndex) => {
    const color = colors[seriesIndex % colors.length];
    const points = item.scores.map((score, index) => `${xFor(index).toFixed(1)},${yFor(score.points).toFixed(1)}`).join(' ');
    const dots = item.scores.map((score, index) => {
      const x = xFor(index);
      const y = yFor(score.points);
      const label = `${item.driver} · ${rounds[index].label || rounds[index].id} · ${score.points} pts${score.place ? ` · P${score.place}` : ''}`;
      return `<g><circle cx="${x.toFixed(1)}" cy="${y.toFixed(1)}" r="4.5" fill="${color}" /><title>${escapeHtml(label)}</title></g>`;
    }).join('');
    return `<g><polyline fill="none" stroke="${color}" stroke-width="3" stroke-linecap="round" stroke-linejoin="round" points="${points}" />${dots}</g>`;
  }).join('');

  return `
    <div class="compare-trend-chart-wrap">
      <svg class="compare-trend-chart" viewBox="0 0 ${width} ${height}" role="img" aria-label="Season points trend comparison chart">
        <rect x="0" y="0" width="${width}" height="${height}" rx="18" fill="rgba(255,255,255,0.02)" />
        ${grid}
        <line x1="${margin.left}" y1="${(height - margin.bottom).toFixed(1)}" x2="${(width - margin.right).toFixed(1)}" y2="${(height - margin.bottom).toFixed(1)}" stroke="rgba(255,255,255,0.16)" stroke-width="1" />
        ${lines}
        ${xLabels}
      </svg>
    </div>
  `;
}

function attachGlobalHandlers() {
  document.querySelectorAll('[data-view-button]').forEach(button => {
    button.addEventListener('click', () => {
      button.closest('.view-dock')?.classList.remove('active');
      setView(button.dataset.viewButton);
    });
  });
  document.querySelectorAll('[data-class-filter]').forEach(select => {
    select.addEventListener('change', event => {
      state.selectedClass = event.target.value;
      render();
    });
  });
  document.querySelectorAll('[data-view-dock-toggle]').forEach(button => {
    button.addEventListener('click', event => {
      const dock = event.currentTarget.closest('.view-dock');
      if (!dock) return;
      const isOpen = dock.classList.toggle('active');
      event.currentTarget.setAttribute('aria-expanded', String(isOpen));
    });
  });
}


function attachViewDockDismissHandlers() {
  if (window.__opulentViewDockDismissAttached) return;
  window.__opulentViewDockDismissAttached = true;

  document.addEventListener('click', event => {
    const dock = event.target.closest?.('.view-dock');
    document.querySelectorAll('.view-dock.active').forEach(openDock => {
      if (openDock !== dock) {
        openDock.classList.remove('active');
        openDock.querySelector('[data-view-dock-toggle]')?.setAttribute('aria-expanded', 'false');
      }
    });
  });

  document.addEventListener('keydown', event => {
    if (event.key !== 'Escape') return;
    document.querySelectorAll('.view-dock.active').forEach(openDock => {
      openDock.classList.remove('active');
      openDock.querySelector('[data-view-dock-toggle]')?.setAttribute('aria-expanded', 'false');
    });
  });
}

attachViewDockDismissHandlers();

function toggleDiag() {
  document.getElementById('diag')?.classList.toggle('open');
}
window.toggleDiag = toggleDiag;

function updateDiagnostics(rawWorkbook, parsed, fileName) {
  const lines = [
    'MULTI-ROUND WORKBOOK',
    `File: ${fileName}`,
    `Workbook sheets: ${rawWorkbook.sheets.length}`,
    `Parsed rounds: ${parsed.rounds.length}`,
    `Total parsed entries: ${parsed.rounds.reduce((sum, round) => sum + round.entries.length, 0)}`,
    '',
    ...parsed.rounds.map(round => `${round.label}: ${round.entries.length} entries, ${round.classOrder.length} classes, ${round.date || 'no date'}`)
  ];
  const diag = document.getElementById('diagText');
  if (diag) diag.textContent = lines.join('\n');
}

/* -----------------------------
   COMPETITION WORKBOOK PARSER
----------------------------- */

function parseCompetitionWorkbook(rawWorkbook, fileName) {
  const rounds = rawWorkbook.sheets
    .map(sheet => parseRoundSheet(sheet))
    .filter(round => round.entries.length);

  rounds.sort((a, b) => roundSortValue(a) - roundSortValue(b));

  return {
    fileName,
    rounds,
    seasonYear: inferSeasonYear(rounds, fileName),
    sourceSheets: rawWorkbook.sheets.map(sheet => sheet.name)
  };
}

function inferSeasonYear(rounds, fileName) {
  const text = [fileName, ...(rounds || []).flatMap(round => [round.title, round.date, round.label, round.sheetName])].filter(Boolean).join(' ');
  const match = String(text).match(/\b(20\d{2})\b/);
  return match ? match[1] : '';
}

function parseRoundSheet(sheet) {
  const rows = sheet.rows || [];
  const titleText = findTitleText(rows) || sheet.name;
  const meta = parseRoundMeta(titleText, sheet.name);
  const classes = {};
  const classOrder = [];
  const entries = [];
  let currentClass = '';
  let currentIndex = null;

  rows.forEach(row => {
    const first = cleanCell(row[0]);
    const second = cleanCell(row[1]);
    const third = cleanCell(row[2]);

    if (!first) return;
    if (/^pos$/i.test(first) && /^driver$/i.test(second)) return;
    if (/best overall|class positions|best pax/i.test(first)) return;

    const possibleClass = isClassHeaderRow(row);
    if (possibleClass) {
      currentClass = normalizeClassHeader(first);
      currentIndex = parseNumber(row[5]);
      if (!classes[currentClass]) {
        classes[currentClass] = [];
        classOrder.push(currentClass);
      }
      return;
    }

    if (!currentClass) return;

    const position = parseNumber(row[0]);
    const driver = cleanCell(row[1]);
    if (!Number.isFinite(position) || !driver) return;

    const rawNumber = cleanCell(row[2]);
    const number = rawNumber ? rawNumber.replace(/\.0$/, '') : '';
    const car = cleanCell(row[3]);
    const bestRaw = parseNumber(row[5]);
    const bestPax = parseNumber(row[6]);
    const paxClass = cleanCell(row[7]);
    const runObjects = parseRuns(row[8]);
    const displayClass = currentClass;

    const entry = {
      roundId: meta.id,
      roundLabel: meta.label,
      position,
      rank: position,
      driver,
      number,
      car,
      sponsors: cleanCell(row[4]),
      cls: currentClass,
      class: paxClass || currentClass,
      displayClass,
      classNumber: [displayClass, number].filter(Boolean).join(' '),
      indexValue: currentIndex,
      bestRaw: Number.isFinite(bestRaw) ? bestRaw : null,
      bestPax: Number.isFinite(bestPax) ? bestPax : null,
      rawTime: Number.isFinite(bestRaw) ? bestRaw : null,
      indexedTime: Number.isFinite(bestPax) ? bestPax : null,
      runs: runObjects.map(run => run.label),
      runObjects
    };

    classes[currentClass].push(entry);
    entries.push(entry);
  });

  Object.values(classes).forEach(rows => rows.sort((a, b) => a.position - b.position));

  return {
    ...meta,
    sheetName: sheet.name,
    title: titleText,
    entries,
    classes,
    classOrder: classOrder.filter(cls => classes[cls]?.length)
  };
}

function findTitleText(rows) {
  for (const row of rows.slice(0, 6)) {
    for (const cell of row || []) {
      const text = cleanCell(cell);
      if (/championship round|class results/i.test(text)) return text;
    }
  }
  return '';
}

function parseRoundMeta(titleText, sheetName) {
  const text = String(titleText || '').replace(/\r/g, '\n');
  const parts = text.split('\n').map(part => part.trim()).filter(Boolean);
  const joined = parts.join(' · ');
  const roundMatch = joined.match(/Round\s*0*(\d+)/i) || String(sheetName).match(/Round\s*0*(\d+)/i);
  const datePart = parts.find(part => /\d{1,2}[-/]\w{3}[-/]\d{4}|\d{1,2}[-/]\d{1,2}[-/]\d{2,4}/i.test(part)) || '';
  const location = parts.find(part => !/championship|class results|round|\d{4}/i.test(part)) || '';
  const roundNumber = roundMatch ? Number(roundMatch[1]) : null;
  const label = roundNumber ? `Round ${roundNumber}` : sheetName;
  return {
    id: normalizeKey(`${sheetName}-${roundNumber || ''}`) || normalizeKey(sheetName),
    roundNumber,
    label,
    date: datePart,
    location
  };
}

function cleanCell(value) {
  if (value === null || value === undefined) return '';
  return String(value).trim();
}

function normalizeClassHeader(value) {
  return cleanCell(value).toUpperCase().replace(/\s+/g, '');
}

function isClassHeaderRow(row) {
  const first = normalizeClassHeader(row?.[0]);
  if (!first) return false;
  if (/^(POS|DRIVER|BEST|CLASS)$/i.test(first)) return false;
  if (Number.isFinite(Number(first))) return false;

  // SFR sheets include both standard class headers such as AS, DS, DSTL
  // and grouped/regional headers with trailing hyphens such as CAM-, N-, S1-, ST2-.
  // The original parser only accepted all-letter headers, so once it hit DSTL it
  // ignored CAM-/N-/S1- style headers and placed all following rows under DSTL.
  const looksLikeClass = /^[A-Z][A-Z0-9]{0,4}-?$/.test(first);
  if (!looksLikeClass) return false;

  const trailingCells = [row?.[1], row?.[2], row?.[3], row?.[4], row?.[6], row?.[7], row?.[8]]
    .map(cleanCell)
    .filter(Boolean);

  // Class header rows may have an index value in column F, but should not have
  // driver/car/run data in the other result columns.
  return trailingCells.length === 0;
}

function parseNumber(value) {
  if (typeof value === 'number') return value;
  const text = cleanCell(value).replace(/,/g, '');
  if (!text) return NaN;
  const num = Number.parseFloat(text);
  return Number.isFinite(num) ? num : NaN;
}

function parseRuns(value) {
  const text = cleanCell(value);
  if (!text) return [];
  if (/^\d+(\.\d+)?$/.test(text)) return [makeRun(text)];

  const matches = text.match(/\d+(?:\.\d+)?(?:\s+\([^)]+\))*/g) || [];
  return matches.map(makeRun);
}

function makeRun(label) {
  const text = cleanCell(label).replace(/\s+/g, ' ');
  const penalties = Array.from(text.matchAll(/\(([^)]+)\)/g)).map(match => match[1].trim().toUpperCase());
  const rawTime = parseNumber(text);
  const isDnf = penalties.some(p => p.includes('DNF'));
  const isRerun = penalties.some(p => p.includes('RRN') || p.includes('RERUN') || p.includes('RE-RUN'));
  const conePenalty = penalties.map(p => Number(p)).filter(Number.isFinite).reduce((a, b) => a + b, 0);
  return {
    label: text,
    rawTime: Number.isFinite(rawTime) ? rawTime : null,
    penalties,
    isDnf,
    isRerun,
    hasCone: conePenalty > 0,
    conePenalty,
    isClean: !isDnf && !isRerun && conePenalty === 0
  };
}

/* -----------------------------
   LOCAL XLSX READER, ALL SHEETS
----------------------------- */

async function readXlsxAllSheets(arrayBuffer) {
  const zip = parseZip(arrayBuffer);
  const workbookXml = await zip.text('xl/workbook.xml');
  const workbookRelsXml = await zip.text('xl/_rels/workbook.xml.rels');
  const workbookDoc = parseXml(workbookXml);
  const relsDoc = parseXml(workbookRelsXml);
  const sharedStrings = zip.has('xl/sharedStrings.xml') ? parseSharedStrings(parseXml(await zip.text('xl/sharedStrings.xml'))) : [];
  const relationships = new Map(Array.from(relsDoc.querySelectorAll('Relationship')).map(node => [node.getAttribute('Id'), node.getAttribute('Target') || '']));

  const sheets = [];
  for (const sheet of workbookDoc.querySelectorAll('sheet')) {
    const name = sheet.getAttribute('name') || 'Sheet';
    const relId = sheet.getAttribute('r:id');
    const target = relationships.get(relId);
    if (!target) continue;
    const sheetPath = target.startsWith('xl/') ? target.replace(/^\//, '') : `xl/${target.replace(/^\//, '')}`;
    const sheetXml = await zip.text(sheetPath);
    sheets.push({ name, path: sheetPath, rows: parseSheet(parseXml(sheetXml), sharedStrings) });
  }
  return { sheets };
}

function parseXml(text) {
  const doc = new DOMParser().parseFromString(text, 'application/xml');
  const error = doc.querySelector('parsererror');
  if (error) throw new Error('Could not parse workbook XML.');
  return doc;
}

function parseSharedStrings(doc) {
  return Array.from(doc.querySelectorAll('si')).map(si => Array.from(si.querySelectorAll('t')).map(t => t.textContent || '').join(''));
}

function parseSheet(doc, sharedStrings) {
  const rows = [];
  for (const rowNode of doc.querySelectorAll('sheetData row')) {
    const rowIndex = Number(rowNode.getAttribute('r')) || rows.length + 1;
    const row = [];
    for (const cell of rowNode.querySelectorAll('c')) {
      const ref = cell.getAttribute('r') || '';
      const colIndex = columnIndex(ref.replace(/\d+/g, ''));
      row[colIndex] = readCell(cell, sharedStrings);
    }
    rows[rowIndex - 1] = row;
  }
  return rows.map(row => row || []);
}

function readCell(cell, sharedStrings) {
  const type = cell.getAttribute('t');
  if (type === 'inlineStr') return cell.querySelector('is t')?.textContent || '';
  if (type === 'str') return cell.querySelector('v')?.textContent || '';
  const value = cell.querySelector('v')?.textContent;
  if (value === undefined || value === null) return '';
  if (type === 's') return sharedStrings[Number(value)] || '';
  if (type === 'b') return value === '1';
  const numeric = Number(value);
  return Number.isFinite(numeric) ? numeric : value;
}

function columnIndex(letters) {
  let index = 0;
  for (const char of letters) index = index * 26 + char.charCodeAt(0) - 64;
  return Math.max(0, index - 1);
}

function parseZip(arrayBuffer) {
  const bytes = new Uint8Array(arrayBuffer);
  const view = new DataView(arrayBuffer);
  const eocdOffset = findEndOfCentralDirectory(view);
  const entries = readCentralDirectory(view, bytes, eocdOffset);

  return {
    has(name) { return entries.has(name); },
    async text(name) {
      const entry = entries.get(name);
      if (!entry) throw new Error(`Missing workbook part: ${name}`);
      const data = readEntryData(view, bytes, entry);
      const inflated = await inflateEntry(data, entry.compressionMethod);
      return new TextDecoder('utf-8').decode(inflated);
    }
  };
}

function findEndOfCentralDirectory(view) {
  for (let offset = view.byteLength - 22; offset >= Math.max(0, view.byteLength - 66000); offset--) {
    if (view.getUint32(offset, true) === 0x06054b50) return offset;
  }
  throw new Error('This does not look like a valid .xlsx ZIP file.');
}

function readCentralDirectory(view, bytes, eocdOffset) {
  const entryCount = view.getUint16(eocdOffset + 10, true);
  let offset = view.getUint32(eocdOffset + 16, true);
  const entries = new Map();

  for (let i = 0; i < entryCount; i++) {
    if (view.getUint32(offset, true) !== 0x02014b50) throw new Error('Invalid ZIP central directory.');
    const compressionMethod = view.getUint16(offset + 10, true);
    const compressedSize = view.getUint32(offset + 20, true);
    const uncompressedSize = view.getUint32(offset + 24, true);
    const nameLength = view.getUint16(offset + 28, true);
    const extraLength = view.getUint16(offset + 30, true);
    const commentLength = view.getUint16(offset + 32, true);
    const localHeaderOffset = view.getUint32(offset + 42, true);
    const nameBytes = bytes.slice(offset + 46, offset + 46 + nameLength);
    const name = new TextDecoder('utf-8').decode(nameBytes);
    entries.set(name, { compressionMethod, compressedSize, uncompressedSize, localHeaderOffset });
    offset += 46 + nameLength + extraLength + commentLength;
  }
  return entries;
}

function readEntryData(view, bytes, entry) {
  const offset = entry.localHeaderOffset;
  if (view.getUint32(offset, true) !== 0x04034b50) throw new Error('Invalid ZIP local file header.');
  const nameLength = view.getUint16(offset + 26, true);
  const extraLength = view.getUint16(offset + 28, true);
  const dataStart = offset + 30 + nameLength + extraLength;
  return bytes.slice(dataStart, dataStart + entry.compressedSize);
}

async function inflateEntry(data, method) {
  if (method === 0) return data;
  if (method !== 8) throw new Error(`Unsupported ZIP compression method: ${method}`);
  if (!('DecompressionStream' in window)) {
    throw new Error('This browser cannot decompress local .xlsx files. Use current Chrome or Edge.');
  }
  const stream = new Blob([data]).stream().pipeThrough(new DecompressionStream('deflate-raw'));
  const buffer = await new Response(stream).arrayBuffer();
  return new Uint8Array(buffer);
}
