// ─────────────────────────────────────────────
// PCms Quality Tracker — app.js
// ─────────────────────────────────────────────

'use strict';

// ── State ────────────────────────────────────
let rawRows = [];       // parsed flat rows from Excel
let koRows  = [];       // KO-only rows
let charts  = {};       // chart instances

// Category columns in the Summary sheet (columns after Ticket ID)
const CATEGORY_COLS = [
  '1. Valid Rejection',
  '2. Communication',
  '3. COMPASS Solution',
  '4. Solution Linkage',
  '5. Related Incident',
  '6. Translation',
  '7. State and State Reason',
  '8. Documentation Traceability',
  '9. Reasonable Ticket Processing',
  '10. Service Efficiency',
  '11. Other Errors'
];

// ── Tab navigation ───────────────────────────
document.querySelectorAll('.nav-btn').forEach(btn => {
  btn.addEventListener('click', () => {
    document.querySelectorAll('.nav-btn').forEach(b => b.classList.remove('active'));
    document.querySelectorAll('.tab').forEach(t => t.classList.add('hidden'));
    btn.classList.add('active');
    document.getElementById('tab-' + btn.dataset.tab).classList.remove('hidden');
  });
});

// ── File upload ──────────────────────────────
const dropZone  = document.getElementById('dropZone');
const fileInput = document.getElementById('fileInput');

dropZone.addEventListener('click', () => fileInput.click());
dropZone.addEventListener('dragover', e => { e.preventDefault(); dropZone.classList.add('drag-over'); });
dropZone.addEventListener('dragleave', () => dropZone.classList.remove('drag-over'));
dropZone.addEventListener('drop', e => {
  e.preventDefault();
  dropZone.classList.remove('drag-over');
  handleFile(e.dataTransfer.files[0]);
});
fileInput.addEventListener('change', e => handleFile(e.target.files[0]));

function handleFile(file) {
  if (!file) return;
  const reader = new FileReader();
  reader.onload = e => {
    try {
      const wb = XLSX.read(e.target.result, { type: 'array' });
      parseWorkbook(wb);
    } catch (err) {
      showError('Failed to read file: ' + err.message);
    }
  };
  reader.readAsArrayBuffer(file);
}

// ── Core parser ──────────────────────────────
function parseWorkbook(wb) {
  // Find the Summary sheet (case-insensitive)
  const sheetName = wb.SheetNames.find(n =>
    n.toLowerCase().includes('summary') || n.toLowerCase().includes('overall')
  ) || wb.SheetNames[0];

  const ws   = wb.Sheets[sheetName];
  const data = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });

  if (data.length < 2) { showError('Sheet appears empty.'); return; }

  // Detect header row: must contain 'Ticket' or 'INC'
  let headerIdx = 0;
  for (let i = 0; i < Math.min(10, data.length); i++) {
    const row = data[i].map(c => String(c).toLowerCase());
    if (row.some(c => c.includes('ticket') || c.includes('inc') || c.includes('week'))) {
      headerIdx = i;
      break;
    }
  }

  const headers = data[headerIdx].map(h => String(h).trim());

  // Locate key columns
  const col = name => headers.findIndex(h => h.toLowerCase().includes(name.toLowerCase()));
  const weekCol   = col('week');
  const monthCol  = col('month');
  const ticketCol = headers.findIndex(h =>
    h.toLowerCase().includes('ticket') || h.toLowerCase().includes('inc')
  );

  if (ticketCol === -1) { showError('Cannot find Ticket ID column. Check the sheet header row.'); return; }

  // Detect category columns: find pairs of (error description, agent)
  // Pattern: columns after ticketCol in groups: [cat1_reason, cat1_agent, cat2_reason, cat2_agent, ...]
  // OR: header contains the category name and next header is blank/"Agent"
  const catColMap = detectCategoryColumns(headers, ticketCol);

  rawRows = [];

  for (let i = headerIdx + 1; i < data.length; i++) {
    const row = data[i];
    if (!row || row.every(c => c === '')) continue;

    const week   = weekCol  >= 0 ? String(row[weekCol]  || '').trim() : '';
    const month  = monthCol >= 0 ? String(row[monthCol] || '').trim() : '';
    const ticket = String(row[ticketCol] || '').trim();

    if (!ticket) continue;

    // For each category column pair found
    for (const { catNum, catLabel, reasonColIdx, agentColIdx } of catColMap) {
      const reasonRaw = reasonColIdx >= 0 ? String(row[reasonColIdx] || '').trim() : '';
      const agentRaw  = agentColIdx  >= 0 ? String(row[agentColIdx]  || '').trim() : '';

      if (!reasonRaw && !agentRaw) continue;

      // Split multiple agents if comma-separated in agent cell
      const agents = agentRaw ? agentRaw.split(/[,;]/).map(a => a.trim()).filter(Boolean) : ['Unknown'];

      for (const agent of agents) {
        const isKO  = reasonRaw.toUpperCase().includes('KO');
        const isNOK = reasonRaw.toUpperCase().includes('NOK');
        if (!isKO && !isNOK && reasonRaw) {
          // might just be a description without prefix — still log it
        }

        rawRows.push({
          week, month, ticket,
          catNum, catLabel,
          reason: reasonRaw,
          agent,
          isKO,
          isNOK
        });
      }
    }
  }

  if (rawRows.length === 0) { showError('No breach data found. Ensure the Summary sheet has the expected format.'); return; }

  // Apply dedup rule: same ticket + same agent = count as 1 breach
  markDuplicates();

  // Build KO-only set
  koRows = rawRows.filter(r => r.isKO);

  // Populate filters
  populateFilters();

  // Show status
  showStatus(`✅ Loaded ${rawRows.length} breach records (${koRows.length} KOs) from "${sheetName}"`);

  // Switch to overview
  document.querySelector('[data-tab="overview"]').click();
  renderOverview();
  renderKOTable();
}

// ── Category column detection ─────────────────
function detectCategoryColumns(headers, ticketCol) {
  const result = [];

  // Strategy 1: headers explicitly contain the category numbers (1., 2., … 11.)
  // Each category may span 2 cols: [reason_col, agent_col] where agent col header has "Agent"
  // or the header is blank
  const catPattern = /^(\d{1,2})[\.\s]/;

  for (let i = ticketCol + 1; i < headers.length; i++) {
    const h = headers[i];
    const match = catPattern.exec(h);
    if (match) {
      const catNum = parseInt(match[1]);
      const catLabel = CATEGORY_COLS[catNum - 1] || h;
      // Next column is likely the agent column (blank header or says 'agent')
      const nextH = (headers[i + 1] || '').toLowerCase();
      const agentIdx = (nextH === '' || nextH.includes('agent') || nextH.includes('name')) ? i + 1 : -1;
      result.push({ catNum, catLabel, reasonColIdx: i, agentColIdx: agentIdx });
      if (agentIdx > 0) i++; // skip the agent col
    }
  }

  // Strategy 2: if nothing found, assume paired columns starting after ticket
  // Each pair = (reason, agent) repeated 11 times
  if (result.length === 0) {
    let colIdx = ticketCol + 1;
    for (let c = 0; c < 11 && colIdx < headers.length; c++) {
      result.push({
        catNum: c + 1,
        catLabel: CATEGORY_COLS[c],
        reasonColIdx: colIdx,
        agentColIdx: colIdx + 1 < headers.length ? colIdx + 1 : -1
      });
      colIdx += 2;
    }
  }

  return result;
}

// ── Dedup marking ─────────────────────────────
function markDuplicates() {
  // Rule: same ticket + same agent → only first occurrence is "primary" (counted)
  // We mark each row with .isDuplicate = true/false
  // Also flag multi-agent breaches on same ticket
  const seen = new Map();

  // Sort by week then ticket so first occurrence is deterministic
  rawRows.sort((a, b) => {
    const wa = parseInt(a.week) || 0, wb_ = parseInt(b.week) || 0;
    if (wa !== wb_) return wa - wb_;
    return a.ticket.localeCompare(b.ticket);
  });

  for (const row of rawRows) {
    const key = `${row.ticket}||${row.agent}||${row.catNum}`;
    if (seen.has(key)) {
      row.isDuplicate = true;
    } else {
      seen.set(key, true);
      row.isDuplicate = false;
    }
  }

  // Flag tickets breached by multiple agents
  const ticketAgents = new Map();
  for (const row of rawRows) {
    if (!row.isDuplicate) {
      if (!ticketAgents.has(row.ticket)) ticketAgents.set(row.ticket, new Set());
      ticketAgents.get(row.ticket).add(row.agent);
    }
  }
  for (const row of rawRows) {
    const agents = ticketAgents.get(row.ticket);
    row.isMultiAgent = agents && agents.size > 1;
  }
}

// ── Filters ───────────────────────────────────
function populateFilters() {
  const months  = [...new Set(rawRows.map(r => r.month ).filter(Boolean))].sort();
  const weeks   = [...new Set(rawRows.map(r => r.week  ).filter(Boolean))].sort((a,b)=>parseInt(a)-parseInt(b));
  const agents  = [...new Set(rawRows.map(r => r.agent ).filter(Boolean))].sort();
  const cats    = [...new Set(rawRows.map(r => r.catLabel).filter(Boolean))].sort();

  fillSelect('filterMonth',        months,  'All Months');
  fillSelect('filterWeek',         weeks,   'All Weeks');
  fillSelect('filterAgentOverview',agents,  'All Agents');
  fillSelect('koFilterMonth',      months,  'All Months');
  fillSelect('koFilterWeek',       weeks,   'All Weeks');
  fillSelect('koFilterAgent',      agents,  'All Agents');
  fillSelect('koFilterCategory',   cats,    'All Categories');

  ['filterMonth','filterWeek','filterAgentOverview'].forEach(id =>
    document.getElementById(id).addEventListener('change', renderOverview));
  ['koFilterMonth','koFilterWeek','koFilterAgent','koFilterCategory'].forEach(id =>
    document.getElementById(id).addEventListener('change', renderKOTable));
  document.getElementById('koSearch').addEventListener('input', renderKOTable);
}

function fillSelect(id, values, placeholder) {
  const el = document.getElementById(id);
  el.innerHTML = `<option value="">${placeholder}</option>` +
    values.map(v => `<option value="${v}">${v}</option>`).join('');
}

function getFilteredOverviewRows() {
  const month = document.getElementById('filterMonth').value;
  const week  = document.getElementById('filterWeek').value;
  const agent = document.getElementById('filterAgentOverview').value;
  return rawRows.filter(r =>
    (!r.isDuplicate) &&
    (!month || r.month === month) &&
    (!week  || r.week  === week ) &&
    (!agent || r.agent === agent)
  );
}

function getFilteredKORows() {
  const month  = document.getElementById('koFilterMonth').value;
  const week   = document.getElementById('koFilterWeek').value;
  const agent  = document.getElementById('koFilterAgent').value;
  const cat    = document.getElementById('koFilterCategory').value;
  const search = document.getElementById('koSearch').value.toLowerCase();
  return koRows.filter(r =>
    (!r.isDuplicate) &&
    (!month  || r.month    === month ) &&
    (!week   || r.week     === week  ) &&
    (!agent  || r.agent    === agent ) &&
    (!cat    || r.catLabel === cat   ) &&
    (!search || r.ticket.toLowerCase().includes(search))
  );
}

// ── Overview rendering ────────────────────────
function renderOverview() {
  const rows  = getFilteredOverviewRows();
  const koR   = rows.filter(r => r.isKO);
  const nokR  = rows.filter(r => r.isNOK);

  const tickets = new Set(koR.map(r => r.ticket)).size;

  // Stats
  document.getElementById('statsRow').innerHTML = `
    <div class="stat-card red"><div class="label">KO Breaches</div><div class="value">${koR.length}</div></div>
    <div class="stat-card warn"><div class="label">NOK Breaches</div><div class="value">${nokR.length}</div></div>
    <div class="stat-card blue"><div class="label">Unique Tickets (KO)</div><div class="value">${tickets}</div></div>
    <div class="stat-card green"><div class="label">Total Records</div><div class="value">${rows.length}</div></div>
  `;

  buildChartByMonth(koR);
  buildChartByAgent(koR);
  buildChartByCategory(koR);
  buildChartWeekly(koR);
}

// Sort months chronologically
function sortMonthsChron(months) {
  const ORDER = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];
  return months.sort((a, b) => {
    const ia = ORDER.findIndex(m => a.includes(m));
    const ib = ORDER.findIndex(m => b.includes(m));
    if (ia !== ib) return ia - ib;
    // year part
    const ya = parseInt(a.replace(/\D/g,'')) || 0;
    const yb = parseInt(b.replace(/\D/g,'')) || 0;
    return ya - yb;
  });
}

function countBy(rows, key) {
  return rows.reduce((acc, r) => {
    const k = r[key] || 'Unknown';
    acc[k] = (acc[k] || 0) + 1;
    return acc;
  }, {});
}

const PALETTE = [
  '#5b6ef5','#e05c5c','#4ecdc4','#f5a623','#a29bfe',
  '#fd79a8','#00b894','#fdcb6e','#74b9ff','#e17055',
  '#55efc4','#fab1a0'
];

function destroyChart(id) {
  if (charts[id]) { charts[id].destroy(); delete charts[id]; }
}

function buildChartByMonth(rows) {
  destroyChart('month');
  const counts = countBy(rows, 'month');
  const labels = sortMonthsChron(Object.keys(counts));
  const values = labels.map(l => counts[l]);
  charts['month'] = new Chart(document.getElementById('chartByMonth'), {
    type: 'bar',
    data: { labels, datasets: [{ label: 'KO Breaches', data: values, backgroundColor: PALETTE[0], borderRadius: 6 }] },
    options: chartOpts('KO Breaches')
  });
}

function buildChartByAgent(rows) {
  destroyChart('agent');
  const counts = countBy(rows, 'agent');
  const sorted = Object.entries(counts).sort((a,b)=>b[1]-a[1]).slice(0,15);
  const labels = sorted.map(e=>e[0]);
  const values = sorted.map(e=>e[1]);
  charts['agent'] = new Chart(document.getElementById('chartByAgent'), {
    type: 'bar',
    data: { labels, datasets: [{ label: 'KO Breaches', data: values, backgroundColor: PALETTE[1], borderRadius: 6 }] },
    options: { ...chartOpts('KO Breaches'), indexAxis: 'y' }
  });
}

function buildChartByCategory(rows) {
  destroyChart('cat');
  const counts = countBy(rows, 'catLabel');
  const sorted = Object.entries(counts).sort((a,b)=>b[1]-a[1]);
  const labels = sorted.map(e=>e[0].replace(/^\d+\.\s*/,''));
  const values = sorted.map(e=>e[1]);
  charts['cat'] = new Chart(document.getElementById('chartByCategory'), {
    type: 'bar',
    data: { labels, datasets: [{ label: 'KO Count', data: values, backgroundColor: PALETTE.slice(0,labels.length), borderRadius: 6 }] },
    options: chartOpts('KO by Category')
  });
}

function buildChartWeekly(rows) {
  destroyChart('weekly');
  const counts = countBy(rows, 'week');
  const labels = Object.keys(counts).sort((a,b)=>parseInt(a)-parseInt(b));
  const values = labels.map(l=>counts[l]);
  charts['weekly'] = new Chart(document.getElementById('chartWeekly'), {
    type: 'line',
    data: {
      labels,
      datasets: [{
        label: 'KO Breaches',
        data: values,
        borderColor: PALETTE[2],
        backgroundColor: 'rgba(78,205,196,.15)',
        fill: true, tension: .4, pointRadius: 4
      }]
    },
    options: chartOpts('KO per Week')
  });
}

function chartOpts(title) {
  return {
    responsive: true,
    plugins: {
      legend: { display: false },
      tooltip: { callbacks: { label: ctx => ` ${ctx.parsed.y ?? ctx.parsed.x} breaches` } }
    },
    scales: {
      x: { ticks: { color: '#8b8fa8', font: { size: 11 } }, grid: { color: '#2e3250' } },
      y: { ticks: { color: '#8b8fa8', font: { size: 11 } }, grid: { color: '#2e3250' } }
    }
  };
}

// ── KO Table rendering ────────────────────────
function renderKOTable() {
  const rows  = getFilteredKORows();
  const tbody = document.getElementById('koTableBody');
  document.getElementById('koCount').textContent = `${rows.length} KO breach${rows.length !== 1 ? 'es' : ''}`;

  if (rows.length === 0) {
    tbody.innerHTML = '<tr><td colspan="7" style="text-align:center;padding:40px;color:var(--muted)">No KO breaches match the current filters.</td></tr>';
    return;
  }

  tbody.innerHTML = rows.map(r => `
    <tr>
      <td>${r.month}</td>
      <td>W${r.week}</td>
      <td style="font-family:monospace;font-size:12px">${r.ticket}</td>
      <td><span class="cat-pill">${r.catLabel.replace(/^\d+\.\s*/,'')}</span></td>
      <td>${r.reason}</td>
      <td><strong>${r.agent}</strong></td>
      <td>${r.isMultiAgent
        ? '<span class="badge badge-multi">Multi-agent</span>'
        : '<span class="badge badge-ok">Single</span>'}
      </td>
    </tr>
  `).join('');
}

// ── Export ────────────────────────────────────
document.getElementById('exportCSV').addEventListener('click', () => {
  const rows = getFilteredKORows();
  if (!rows.length) return;
  const header = 'Month,Week,Ticket ID,Category,KO Description,Agent,Multi-breach Flag\n';
  const body   = rows.map(r =>
    [r.month, `W${r.week}`, r.ticket, r.catLabel, `"${r.reason}"`, r.agent,
     r.isMultiAgent ? 'Multi-agent' : 'Single'].join(',')
  ).join('\n');
  download('PCms_KO_Breaches.csv', 'text/csv', header + body);
});

document.getElementById('exportExcel').addEventListener('click', () => {
  const rows = getFilteredKORows();
  if (!rows.length) return;
  const ws = XLSX.utils.json_to_sheet(rows.map(r => ({
    Month:            r.month,
    Week:             `W${r.week}`,
    'Ticket ID':      r.ticket,
    Category:         r.catLabel,
    'KO Description': r.reason,
    Agent:            r.agent,
    'Multi-breach Flag': r.isMultiAgent ? 'Multi-agent' : 'Single'
  })));
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, 'KO Breaches');
  XLSX.writeFile(wb, 'PCms_KO_Breaches.xlsx');
});

function download(filename, mime, content) {
  const a   = document.createElement('a');
  a.href    = URL.createObjectURL(new Blob([content], { type: mime }));
  a.download = filename;
  a.click();
}

// ── UI helpers ────────────────────────────────
function showStatus(msg) {
  const el = document.getElementById('uploadStatus');
  el.textContent = msg;
  el.classList.remove('hidden');
  document.getElementById('uploadError').classList.add('hidden');
}
function showError(msg) {
  const el = document.getElementById('uploadError');
  el.textContent = msg;
  el.classList.remove('hidden');
  document.getElementById('uploadStatus').classList.add('hidden');
}
