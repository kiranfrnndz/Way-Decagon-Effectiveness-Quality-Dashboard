/* ===================================================
   WAY-DECAGON DASHBOARD — APP.JS
   Version: 2.0.0 | Production Ready
=================================================== */

'use strict';

// ══════════════════════════════════════════════════
// CONFIGURATION
// ══════════════════════════════════════════════════
const CONFIG = {
  AI_INTERACTION_TYPE: 'AI-Agent Call',
  INTERNAL_TYPES: ['TL Review','Supervisor Review','Manager Review','Internal Notes','Internal Follow Up','Vendor Follow Up','Quality Review','QC Audit','Internal'],
  CUSTOMER_FACING: ['AI-Agent Call','Call','Email','Chat','SMS','Chat Message','Inbound Call','Outbound Call'],
  COMPLIANCE_FIELDS: ['Reason','Sub Reason','Action Taken'],
  DEFECT_THRESHOLD_SEC: 60,
  REQUIRED_COLUMNS: ['Ticket ID','Interaction ID','Interaction Type','Created Date'],
  DATE_FORMAT: 'Created Date',
  GOOGLE_SHEETS_ENABLED: false,
  VERSION: '2.0.0'
};

// Column aliases — maps various CSV header names to canonical names
const COLUMN_ALIASES = {
  'ticket id':       'Ticket ID',
  'ticket_id':       'Ticket ID',
  'ticketid':        'Ticket ID',
  'ogi':             'OGI',
  'order group id':  'OGI',
  'interaction id':  'Interaction ID',
  'interaction_id':  'Interaction ID',
  'interaction type':'Interaction Type',
  'channel':         'Interaction Type',
  'type':            'Interaction Type',
  'created date':    'Created Date',
  'created_date':    'Created Date',
  'date':            'Created Date',
  'created at':      'Created Date',
  'reason':          'Reason',
  'sub reason':      'Sub Reason',
  'sub_reason':      'Sub Reason',
  'subreason':       'Sub Reason',
  'action taken':    'Action Taken',
  'action_taken':    'Action Taken',
  'resolution':      'Action Taken',
  'status':          'Status',
  'assignee':        'Assignee'
};

// ══════════════════════════════════════════════════
// STATE
// ══════════════════════════════════════════════════
const STATE = {
  rawRows: [],
  normalizedRows: [],
  ticketMap: new Map(),
  filteredTickets: new Map(),
  charts: {},
  datatables: {},
  validationPassed: false,
  filterActive: false,
  currentTheme: 'dark'
};

// ══════════════════════════════════════════════════
// UTILITY FUNCTIONS
// ══════════════════════════════════════════════════
const fmt = {
  num: n => n == null ? '—' : Number(n).toLocaleString(),
  pct: n => n == null ? '—' : Number(n).toFixed(1) + '%',
  date: d => {
    if (!d) return '—';
    const dt = new Date(d);
    return isNaN(dt) ? String(d) : dt.toLocaleDateString('en-US',{month:'short',day:'numeric',year:'numeric'});
  },
  time: d => {
    if (!d) return '—';
    const dt = new Date(d);
    return isNaN(dt) ? String(d) : dt.toLocaleTimeString('en-US',{hour:'2-digit',minute:'2-digit',second:'2-digit'});
  },
  datetime: d => {
    if (!d) return '—';
    const dt = new Date(d);
    return isNaN(dt) ? String(d) : dt.toLocaleString('en-US',{month:'short',day:'numeric',year:'numeric',hour:'2-digit',minute:'2-digit'});
  }
};

function showToast(msg, type='info', duration=4000) {
  const tc = document.getElementById('toastContainer');
  const t = document.createElement('div');
  const icons = {success:'fa-circle-check', error:'fa-circle-xmark', info:'fa-circle-info'};
  t.className = `toast toast-${type}`;
  t.innerHTML = `<i class="fa-solid ${icons[type]||icons.info}"></i>${msg}`;
  tc.appendChild(t);
  setTimeout(() => { t.style.animation='fadeOut 0.3s ease forwards'; setTimeout(()=>t.remove(),300); }, duration);
}

function unlockSection(id) {
  document.getElementById(id)?.classList.add('unlocked');
}

function pct(num, den) { return den > 0 ? (num / den * 100) : 0; }

function badge(text, color='muted') {
  return `<span class="badge badge-${color}">${text}</span>`;
}

function colorPct(p) {
  if (p >= 80) return 'badge-green';
  if (p >= 60) return 'badge-amber';
  return 'badge-red';
}

// ══════════════════════════════════════════════════
// COLUMN NORMALIZATION
// ══════════════════════════════════════════════════
function normalizeColumns(rows) {
  if (!rows.length) return rows;
  const headers = Object.keys(rows[0]);
  const mapping = {};
  headers.forEach(h => {
    const lower = h.trim().toLowerCase();
    if (COLUMN_ALIASES[lower]) mapping[h] = COLUMN_ALIASES[lower];
    else mapping[h] = h.trim(); // keep original if no alias
  });
  return rows.map(row => {
    const norm = {};
    Object.entries(row).forEach(([k, v]) => {
      norm[mapping[k] || k] = v;
    });
    return norm;
  });
}

// ══════════════════════════════════════════════════
// VALIDATION
// ══════════════════════════════════════════════════
function validateData(rows) {
  if (!rows.length) return { passed: false, checks: {}, counts: {} };
  const cols = Object.keys(rows[0]);

  const checks = {
    cols: CONFIG.REQUIRED_COLUMNS.some(c => cols.includes(c)),
    ticket: rows.some(r => r['Ticket ID'] && String(r['Ticket ID']).trim()),
    ogi: rows.some(r => r['OGI'] && String(r['OGI']).trim()),
    types: rows.some(r => r['Interaction Type'] && String(r['Interaction Type']).trim()),
    compliance: CONFIG.COMPLIANCE_FIELDS.some(f => cols.includes(f))
  };

  const counts = {
    records: rows.length,
    uniqueOGI: new Set(rows.map(r => r['OGI']).filter(Boolean)).size,
    uniqueTickets: new Set(rows.map(r => r['Ticket ID']).filter(Boolean)).size,
    totalInteractions: rows.length,
    aiInteractions: rows.filter(r => String(r['Interaction Type']||'').trim() === CONFIG.AI_INTERACTION_TYPE).length,
    humanInteractions: rows.filter(r => {
      const t = String(r['Interaction Type']||'').trim();
      return CONFIG.CUSTOMER_FACING.includes(t) && t !== CONFIG.AI_INTERACTION_TYPE;
    }).length,
    internalInteractions: rows.filter(r => CONFIG.INTERNAL_TYPES.includes(String(r['Interaction Type']||'').trim())).length
  };

  const passed = checks.cols && checks.ticket && checks.types;
  return { passed, checks, counts };
}

// ══════════════════════════════════════════════════
// TICKET MAP BUILDER
// ══════════════════════════════════════════════════
function buildTicketMap(rows) {
  const map = new Map();

  rows.forEach(row => {
    const ticketId = String(row['Ticket ID'] || '').trim();
    if (!ticketId) return;

    if (!map.has(ticketId)) {
      map.set(ticketId, {
        ticketId,
        ogi: String(row['OGI'] || '').trim() || 'UNKNOWN',
        createdDate: row['Created Date'] || '',
        reason: String(row['Reason'] || '').trim(),
        subReason: String(row['Sub Reason'] || '').trim(),
        actionTaken: String(row['Action Taken'] || '').trim(),
        interactions: [],
        status: String(row['Status'] || '').trim(),
        assignee: String(row['Assignee'] || '').trim()
      });
    }

    const ticket = map.get(ticketId);
    ticket.interactions.push({
      interactionId: String(row['Interaction ID'] || '').trim(),
      type: String(row['Interaction Type'] || '').trim(),
      createdDate: row['Created Date'] || '',
      parsedDate: parseDate(row['Created Date']),
      reason: String(row['Reason'] || '').trim(),
      subReason: String(row['Sub Reason'] || '').trim(),
      actionTaken: String(row['Action Taken'] || '').trim(),
      raw: row
    });

    // Update ticket-level fields from latest interaction if blank
    if (!ticket.reason && row['Reason']) ticket.reason = String(row['Reason']).trim();
    if (!ticket.subReason && row['Sub Reason']) ticket.subReason = String(row['Sub Reason']).trim();
    if (!ticket.actionTaken && row['Action Taken']) ticket.actionTaken = String(row['Action Taken']).trim();
    if (!ticket.ogi || ticket.ogi === 'UNKNOWN') ticket.ogi = String(row['OGI'] || '').trim() || 'UNKNOWN';
  });

  // Sort interactions by date
  map.forEach(ticket => {
    ticket.interactions.sort((a, b) => (a.parsedDate||0) - (b.parsedDate||0));
    // Use earliest interaction date as ticket created date
    if (ticket.interactions.length && ticket.interactions[0].parsedDate) {
      ticket.createdDate = ticket.interactions[0].createdDate;
    }
    // Derive all computed metrics
    enrichTicket(ticket);
  });

  return map;
}

function parseDate(d) {
  if (!d) return null;
  const dt = new Date(d);
  return isNaN(dt) ? null : dt.getTime();
}

function enrichTicket(ticket) {
  const ints = ticket.interactions;

  // classify interactions
  const customerFacing = ints.filter(i => CONFIG.CUSTOMER_FACING.includes(i.type));
  const internalInts = ints.filter(i => CONFIG.INTERNAL_TYPES.includes(i.type));
  const aiInts = ints.filter(i => i.type === CONFIG.AI_INTERACTION_TYPE);
  const humanCustomerInts = customerFacing.filter(i => i.type !== CONFIG.AI_INTERACTION_TYPE);

  ticket.isAITicket = aiInts.length > 0;
  ticket.aiInteractionCount = aiInts.length;
  ticket.humanInteractionCount = humanCustomerInts.length;
  ticket.internalInteractionCount = internalInts.length;
  ticket.customerFacingCount = customerFacing.length;

  // FCR: only customer-facing count matters
  ticket.fcrAchieved = customerFacing.length <= 1;

  // AI Containment: AI ticket with no human customer-facing after AI
  if (ticket.isAITicket) {
    const firstAI = aiInts[0];
    const humanAfterAI = humanCustomerInts.filter(i => (i.parsedDate||0) > (firstAI.parsedDate||0));
    ticket.aiContained = humanAfterAI.length === 0;
    ticket.humanTouchAfterAI = humanAfterAI.length > 0;
    ticket.escalated = humanAfterAI.some(i => ['Call','Inbound Call','Outbound Call'].includes(i.type));
  } else {
    ticket.aiContained = false;
    ticket.humanTouchAfterAI = false;
    ticket.escalated = false;
  }

  // Compliance (only for AI tickets)
  if (ticket.isAITicket) {
    ticket.missingReason = !ticket.reason;
    ticket.missingSubReason = !ticket.subReason;
    ticket.missingAction = !ticket.actionTaken;
    ticket.compliant = !ticket.missingReason && !ticket.missingSubReason && !ticket.missingAction;
  } else {
    ticket.missingReason = false;
    ticket.missingSubReason = false;
    ticket.missingAction = false;
    ticket.compliant = true;
  }

  // Defect detection
  ticket.hasDuplicateAI = aiInts.length > 1;
  ticket.duplicateAICount = Math.max(0, aiInts.length - 1);

  // Same-timestamp defects
  const aiTimestamps = aiInts.map(i => i.parsedDate).filter(Boolean);
  const tsSet = new Set();
  let sameTs = 0;
  aiTimestamps.forEach(ts => { if (tsSet.has(ts)) sameTs++; else tsSet.add(ts); });
  ticket.sameTimestampDefects = sameTs;

  // Short interval defects (configurable threshold)
  let shortInterval = 0;
  for (let i = 1; i < aiInts.length; i++) {
    const prev = aiInts[i-1].parsedDate || 0;
    const curr = aiInts[i].parsedDate || 0;
    if (curr > 0 && prev > 0 && (curr - prev) < CONFIG.DEFECT_THRESHOLD_SEC * 1000) {
      shortInterval++;
    }
  }
  ticket.shortIntervalDefects = shortInterval;
  ticket.hasDefect = ticket.hasDuplicateAI || ticket.sameTimestampDefects > 0 || ticket.shortIntervalDefects > 0;

  // Escalation reason
  ticket.escalationReason = ticket.escalated ? (ticket.reason || 'Unknown') : null;

  // Week/date bucket
  if (ticket.createdDate) {
    const d = new Date(ticket.createdDate);
    if (!isNaN(d)) {
      ticket.dateBucket = d.toISOString().slice(0,10);
      ticket.weekBucket = getWeekStart(d).toISOString().slice(0,10);
    }
  }
}

function getWeekStart(d) {
  const dt = new Date(d);
  const day = dt.getDay();
  const diff = dt.getDate() - day + (day === 0 ? -6 : 1);
  return new Date(dt.setDate(diff));
}

// ══════════════════════════════════════════════════
// AGGREGATE METRICS
// ══════════════════════════════════════════════════
function computeMetrics(ticketMap) {
  const tickets = [...ticketMap.values()];
  const aiTickets = tickets.filter(t => t.isAITicket);

  return {
    totalTickets: tickets.length,
    aiTickets: aiTickets.length,
    aiInteractions: tickets.reduce((s,t) => s + t.aiInteractionCount, 0),
    fcrAchieved: aiTickets.filter(t => t.fcrAchieved).length,
    fcrRate: pct(aiTickets.filter(t => t.fcrAchieved).length, aiTickets.length),
    aiContained: aiTickets.filter(t => t.aiContained).length,
    containmentRate: pct(aiTickets.filter(t => t.aiContained).length, aiTickets.length),
    humanTouch: aiTickets.filter(t => t.humanTouchAfterAI).length,
    humanTouchRate: pct(aiTickets.filter(t => t.humanTouchAfterAI).length, aiTickets.length),
    escalated: aiTickets.filter(t => t.escalated).length,
    escalationRate: pct(aiTickets.filter(t => t.escalated).length, aiTickets.length),
    compliant: aiTickets.filter(t => t.compliant).length,
    complianceRate: pct(aiTickets.filter(t => t.compliant).length, aiTickets.length),
    missingReason: aiTickets.filter(t => t.missingReason).length,
    missingSubReason: aiTickets.filter(t => t.missingSubReason).length,
    missingAction: aiTickets.filter(t => t.missingAction).length,
    duplicateAITickets: aiTickets.filter(t => t.hasDuplicateAI).length,
    sameTimestampDefects: aiTickets.reduce((s,t) => s + t.sameTimestampDefects, 0),
    shortIntervalDefects: aiTickets.reduce((s,t) => s + t.shortIntervalDefects, 0),
    defectTickets: aiTickets.filter(t => t.hasDefect).length,
    aiOnly: aiTickets.filter(t => t.aiContained && !t.humanTouchAfterAI).length,
    humanAssisted: aiTickets.filter(t => t.humanTouchAfterAI && !t.escalated).length,
    tickets,
    aiTickets
  };
}

// ══════════════════════════════════════════════════
// SAMPLE DATA GENERATOR
// ══════════════════════════════════════════════════
function generateSampleData(n=500) {
  const reasons = ['Booking Modification','Parking Not Found','Payment Issue','Refund Request','Access Issue','General Inquiry','QR Code Problem','App Technical Issue','Lot Directions','Check-in Assistance'];
  const subReasons = ['Change Date','Change Location','Payment Failed','Duplicate Charge','Gate Won't Open','App Crash','QR Expired','Lot Full','Wrong Address','Pricing Question'];
  const actions = ['Modified Booking','Issued Refund','Escalated to Ops','Provided Directions','Reset QR Code','Collected Feedback','Transferred to Lot','Processed Cancellation','Sent Confirmation','Created Manual Pass'];
  const rows = [];

  for (let i = 0; i < n; i++) {
    const ticketNum = 100000 + Math.floor(Math.random() * 50000);
    const ticketId = `TKT-${ticketNum}`;
    const ogi = `OGI-${Math.floor(ticketNum/3)}`;
    const baseDate = new Date('2026-05-01T00:00:00Z');
    baseDate.setDate(baseDate.getDate() + Math.floor(Math.random() * 38));
    baseDate.setHours(8 + Math.floor(Math.random() * 14), Math.floor(Math.random() * 60));
    const reason = reasons[Math.floor(Math.random()*reasons.length)];
    const subReason = subReasons[Math.floor(Math.random()*subReasons.length)];
    const action = Math.random() > 0.12 ? actions[Math.floor(Math.random()*actions.length)] : '';

    // AI interaction (always present)
    const aiDate = new Date(baseDate);
    rows.push({
      'Ticket ID': ticketId,
      'OGI': ogi,
      'Interaction ID': `INT-${100000+i*3}`,
      'Interaction Type': 'AI-Agent Call',
      'Created Date': aiDate.toISOString(),
      'Reason': Math.random() > 0.08 ? reason : '',
      'Sub Reason': Math.random() > 0.10 ? subReason : '',
      'Action Taken': action,
      'Status': 'Closed',
      'Assignee': 'AI Agent'
    });

    // Duplicate AI (10% chance)
    if (Math.random() < 0.10) {
      const d2 = new Date(aiDate);
      d2.setSeconds(d2.getSeconds() + Math.floor(Math.random() < 0.5 ? 0 : Math.random()*120));
      rows.push({
        'Ticket ID': ticketId,
        'OGI': ogi,
        'Interaction ID': `INT-${100000+i*3+1}`,
        'Interaction Type': 'AI-Agent Call',
        'Created Date': d2.toISOString(),
        'Reason': reason, 'Sub Reason': subReason, 'Action Taken': action,
        'Status': 'Closed', 'Assignee': 'AI Agent'
      });
    }

    // Internal interaction (30% chance)
    if (Math.random() < 0.30) {
      const d3 = new Date(aiDate);
      d3.setMinutes(d3.getMinutes() + 5);
      const internalTypes = ['TL Review','QC Audit','Internal Notes'];
      rows.push({
        'Ticket ID': ticketId,
        'OGI': ogi,
        'Interaction ID': `INT-${100000+i*3+2}`,
        'Interaction Type': internalTypes[Math.floor(Math.random()*internalTypes.length)],
        'Created Date': d3.toISOString(),
        'Reason': reason, 'Sub Reason': subReason, 'Action Taken': action,
        'Status': 'Closed', 'Assignee': 'Team Lead'
      });
    }

    // Human follow-up (25% chance)
    if (Math.random() < 0.25) {
      const d4 = new Date(aiDate);
      d4.setMinutes(d4.getMinutes() + 15 + Math.floor(Math.random()*60));
      const humanTypes = ['Call','Email','Chat'];
      const agents = ['Arya J S','Amal Krishna A','Haleema Raheem','Aswin AV','Nimi M Nair','Ananthu JR'];
      rows.push({
        'Ticket ID': ticketId,
        'OGI': ogi,
        'Interaction ID': `INT-${100000+i*3+3}`,
        'Interaction Type': humanTypes[Math.floor(Math.random()*humanTypes.length)],
        'Created Date': d4.toISOString(),
        'Reason': reason, 'Sub Reason': subReason, 'Action Taken': action,
        'Status': 'Closed',
        'Assignee': agents[Math.floor(Math.random()*agents.length)]
      });
    }
  }

  return rows;
}

// ══════════════════════════════════════════════════
// FILE UPLOAD HANDLING
// ══════════════════════════════════════════════════
function setupUpload() {
  const dropZone = document.getElementById('dropZone');
  const fileInput = document.getElementById('fileInput');

  dropZone.addEventListener('click', () => fileInput.click());
  dropZone.addEventListener('dragover', e => { e.preventDefault(); dropZone.classList.add('drag-over'); });
  dropZone.addEventListener('dragleave', () => dropZone.classList.remove('drag-over'));
  dropZone.addEventListener('drop', e => {
    e.preventDefault();
    dropZone.classList.remove('drag-over');
    const file = e.dataTransfer.files[0];
    if (file) processFile(file);
  });
  fileInput.addEventListener('change', e => {
    if (e.target.files[0]) processFile(e.target.files[0]);
  });

  document.getElementById('loadSampleBtn').addEventListener('click', () => {
    showToast('Generating 500 synthetic records…', 'info');
    setTimeout(() => {
      const rows = generateSampleData(500);
      processRows(rows, 'sample_data_500.csv');
    }, 100);
  });
}

function processFile(file) {
  if (!file.name.endsWith('.csv') && file.type !== 'text/csv') {
    showToast('Please upload a CSV file', 'error'); return;
  }
  showUploadProgress();
  Papa.parse(file, {
    header: true, skipEmptyLines: true, dynamicTyping: false,
    complete: result => {
      if (result.errors.length > 3) {
        showToast('CSV parsing issues — check file format', 'error');
      }
      processRows(result.data, file.name);
    },
    error: err => showToast('Failed to parse CSV: ' + err.message, 'error')
  });
}

function showUploadProgress() {
  const pg = document.getElementById('uploadProgress');
  const bar = document.getElementById('progressBar');
  pg.style.display='block';
  let p=0;
  const iv = setInterval(() => {
    p = Math.min(p + Math.random()*15, 90);
    bar.style.width = p + '%';
    if (p >= 90) clearInterval(iv);
  }, 120);
  STATE._progressInterval = iv;
  STATE._progressBar = bar;
}

function completeProgress() {
  clearInterval(STATE._progressInterval);
  if (STATE._progressBar) STATE._progressBar.style.width = '100%';
  setTimeout(() => { document.getElementById('uploadProgress').style.display='none'; }, 500);
}

function processRows(rows, filename) {
  STATE.rawRows = rows;
  STATE.normalizedRows = normalizeColumns(rows);

  // Show file status
  const now = new Date();
  document.getElementById('fileStatus').style.display='block';
  document.getElementById('statFilename').textContent = filename;
  document.getElementById('statLoaded').textContent = now.toLocaleString('en-GB',{day:'2-digit',month:'short',year:'numeric',hour:'2-digit',minute:'2-digit'}).replace(',','');
  document.getElementById('statRecords').textContent = rows.length.toLocaleString();
  document.getElementById('statStatus').innerHTML = '<span style="color:var(--green)">✓ Processed Successfully</span>';

  completeProgress();

  // Validation
  const val = validateData(STATE.normalizedRows);
  renderValidation(val);

  if (!val.passed) {
    showToast('Validation failed — check required columns', 'error');
    return;
  }

  // Build ticket map
  STATE.ticketMap = buildTicketMap(STATE.normalizedRows);
  STATE.filteredTickets = new Map(STATE.ticketMap);

  // Update hero meta
  const m = computeMetrics(STATE.ticketMap);
  document.getElementById('heroMeta').innerHTML = `
    <span class="meta-pill"><i class="fa-solid fa-database"></i> ${fmt.num(rows.length)} records</span>
    <span class="meta-pill"><i class="fa-solid fa-ticket"></i> ${fmt.num(m.totalTickets)} tickets</span>
    <span class="meta-pill"><i class="fa-solid fa-robot"></i> ${fmt.num(m.aiTickets)} AI tickets</span>
    <span class="meta-pill" style="color:var(--green)"><i class="fa-solid fa-circle-check"></i> Live</span>
  `;

  // Populate filter dropdowns
  populateFilters();

  // Render everything
  renderDashboard();

  showToast(`Loaded ${fmt.num(rows.length)} records — ${fmt.num(m.aiTickets)} AI tickets found`, 'success');
}

// ══════════════════════════════════════════════════
// VALIDATION RENDER
// ══════════════════════════════════════════════════
function renderValidation(val) {
  const { counts, checks, passed } = val;

  document.getElementById('vv-records').textContent = fmt.num(counts.records);
  document.getElementById('vv-ogi').textContent = fmt.num(counts.uniqueOGI);
  document.getElementById('vv-tickets').textContent = fmt.num(counts.uniqueTickets);
  document.getElementById('vv-interactions').textContent = fmt.num(counts.totalInteractions);
  document.getElementById('vv-ai').textContent = fmt.num(counts.aiInteractions);
  document.getElementById('vv-human').textContent = fmt.num(counts.humanInteractions);
  document.getElementById('vv-internal').textContent = fmt.num(counts.internalInteractions);

  const checkMap = {
    'chk-cols':  { ok: checks.cols, label: 'Required columns present' },
    'chk-ticket':{ ok: checks.ticket, label: 'Ticket IDs present' },
    'chk-ogi':   { ok: checks.ogi, label: 'OGI identifiers present' },
    'chk-types': { ok: checks.types, label: 'Interaction types present' },
    'chk-compliance': { ok: checks.compliance, label: 'Compliance fields present' }
  };
  Object.entries(checkMap).forEach(([id, {ok, label}]) => {
    const el = document.getElementById(id);
    el.className = `val-check-item ${ok?'pass':'fail'}`;
    el.innerHTML = `<i class="fa-solid fa-${ok?'circle-check':'circle-xmark'}"></i> ${label}`;
  });

  const badge = document.getElementById('validationBadge');
  badge.className = `validation-badge ${passed?'pass':'fail'}`;
  badge.innerHTML = `<i class="fa-solid fa-${passed?'shield-check':'shield-xmark'}"></i> ${passed?'Validation Passed':'Validation Failed'}`;

  STATE.validationPassed = passed;
  unlockSection('section-validation');
}

// ══════════════════════════════════════════════════
// POPULATE FILTERS
// ══════════════════════════════════════════════════
function populateFilters() {
  const tickets = [...STATE.ticketMap.values()];
  const reasons = [...new Set(tickets.map(t => t.reason).filter(Boolean))].sort();
  const types = [...new Set(STATE.normalizedRows.map(r => r['Interaction Type']).filter(Boolean))].sort();

  const rSelect = document.getElementById('filterReason');
  rSelect.innerHTML = '<option value="">All Reasons</option>';
  reasons.forEach(r => { rSelect.innerHTML += `<option value="${r}">${r}</option>`; });

  const tSelect = document.getElementById('filterType');
  tSelect.innerHTML = '<option value="">All Types</option>';
  types.forEach(t => { tSelect.innerHTML += `<option value="${t}">${t}</option>`; });

  // Date range
  const dates = tickets.map(t => t.createdDate).filter(Boolean).map(d => new Date(d)).filter(d => !isNaN(d));
  if (dates.length) {
    const min = new Date(Math.min(...dates)).toISOString().slice(0,10);
    const max = new Date(Math.max(...dates)).toISOString().slice(0,10);
    document.getElementById('filterDateFrom').value = min;
    document.getElementById('filterDateTo').value = max;
    document.getElementById('filterDateFrom').min = min;
    document.getElementById('filterDateTo').max = max;
  }
}

// ══════════════════════════════════════════════════
// FILTERS
// ══════════════════════════════════════════════════
function applyFilters() {
  const dateFrom = document.getElementById('filterDateFrom').value;
  const dateTo = document.getElementById('filterDateTo').value;
  const reason = document.getElementById('filterReason').value;
  const type = document.getElementById('filterType').value;

  STATE.filteredTickets = new Map();
  STATE.ticketMap.forEach((ticket, id) => {
    if (dateFrom) {
      const d = new Date(ticket.createdDate);
      if (!isNaN(d) && d < new Date(dateFrom)) return;
    }
    if (dateTo) {
      const d = new Date(ticket.createdDate);
      if (!isNaN(d) && d > new Date(dateTo + 'T23:59:59')) return;
    }
    if (reason && ticket.reason !== reason) return;
    if (type && !ticket.interactions.some(i => i.type === type)) return;
    STATE.filteredTickets.set(id, ticket);
  });

  STATE.filterActive = true;
  renderDashboard();
  showToast(`Filter applied — ${fmt.num(STATE.filteredTickets.size)} tickets`, 'info');
}

function clearFilters() {
  STATE.filteredTickets = new Map(STATE.ticketMap);
  STATE.filterActive = false;
  populateFilters();
  renderDashboard();
  showToast('Filters cleared', 'info');
}

// ══════════════════════════════════════════════════
// MAIN DASHBOARD RENDER
// ══════════════════════════════════════════════════
function renderDashboard() {
  const m = computeMetrics(STATE.filteredTickets);

  // Unlock all sections
  ['section-filters','section-kpis','section-effectiveness','section-compliance','section-defects','section-reasons','section-tickets','section-ceo'].forEach(unlockSection);

  renderKPIs(m);
  renderFunnel(m);
  renderEffectivenessCharts(m);
  renderComplianceSection(m);
  renderDefectSection(m);
  renderReasonAnalysis(m);
  renderMasterTable(m);
  renderCEOSummary(m);
}

// ══════════════════════════════════════════════════
// KPI CARDS
// ══════════════════════════════════════════════════
function renderKPIs(m) {
  const kpis = [
    { label:'AI Tickets Created', val:fmt.num(m.aiTickets), pct:null, icon:'fa-robot', color:'cyan', tooltip:'Unique Ticket IDs containing at least one AI-Agent Call interaction', level:'Ticket' },
    { label:'AI Agent Interactions', val:fmt.num(m.aiInteractions), pct:null, icon:'fa-message-bot', color:'purple', tooltip:'Total count of AI-Agent Call interaction records', level:'Interaction' },
    { label:'AI FCR Rate', val:fmt.pct(m.fcrRate), pct:m.fcrRate, icon:'fa-bullseye', color:'green', tooltip:'AI tickets resolved in a single customer-facing interaction (no follow-up needed)', level:'Ticket' },
    { label:'AI Containment Rate', val:fmt.pct(m.containmentRate), pct:m.containmentRate, icon:'fa-shield-halved', color:'green', tooltip:'AI tickets with no customer-facing human interaction after AI', level:'Ticket' },
    { label:'Escalated to Human', val:fmt.num(m.escalated), pct:m.escalationRate, icon:'fa-person-walking-arrow-right', color:'amber', tooltip:'AI tickets where customer was later transferred to a human agent via call', level:'Ticket' },
    { label:'Human Touch Rate', val:fmt.pct(m.humanTouchRate), pct:m.humanTouchRate, icon:'fa-hand-holding', color:'amber', tooltip:'AI tickets that received any customer-facing human contact after AI interaction', level:'Ticket' },
    { label:'Compliance Rate', val:fmt.pct(m.complianceRate), pct:m.complianceRate, icon:'fa-clipboard-check', color:'green', tooltip:'AI tickets with Reason, Sub Reason, and Action Taken all populated', level:'Ticket' },
    { label:'Duplicate AI Tickets', val:fmt.num(m.duplicateAITickets), pct:null, icon:'fa-copy', color:'red', tooltip:'AI tickets containing more than one AI-Agent Call interaction', level:'Ticket' },
    { label:'Same Timestamp Defects', val:fmt.num(m.sameTimestampDefects), pct:null, icon:'fa-clock', color:'red', tooltip:'AI interactions created at identical timestamps within the same ticket', level:'Interaction' },
    { label:'Compliance Failures', val:fmt.num(m.aiTickets - m.compliant), pct:null, icon:'fa-triangle-exclamation', color:'red', tooltip:'AI tickets missing any of: Reason, Sub Reason, or Action Taken', level:'Ticket' }
  ];

  const colorMap = {
    cyan:   { accent: 'var(--cyan)', dim: 'var(--cyan-dim)' },
    purple: { accent: 'var(--purple)', dim: 'var(--purple-dim)' },
    green:  { accent: 'var(--green)', dim: 'var(--green-dim)' },
    amber:  { accent: 'var(--amber)', dim: 'var(--amber-dim)' },
    red:    { accent: 'var(--red)', dim: 'var(--red-dim)' }
  };

  const grid = document.getElementById('kpiGrid');
  grid.innerHTML = kpis.map(k => {
    const c = colorMap[k.color] || colorMap.cyan;
    const pctHtml = k.pct != null ? `<div class="kpi-pct">${fmt.pct(k.pct)}</div>` : '';
    const levelClass = k.level === 'Interaction' ? 'interaction' : '';
    return `
      <div class="kpi-card" style="--accent-color:${c.accent};--accent-dim:${c.dim}">
        <div class="kpi-tooltip" title="${k.tooltip}"><i class="fa-solid fa-circle-info"></i></div>
        <div class="kpi-icon"><i class="fa-solid ${k.icon}"></i></div>
        <div class="kpi-label">${k.label}</div>
        <div class="kpi-value">${k.val}</div>
        ${pctHtml}
        <div class="kpi-level-badge">${k.level} Level</div>
      </div>`;
  }).join('');
}

// ══════════════════════════════════════════════════
// FUNNEL
// ══════════════════════════════════════════════════
function renderFunnel(m) {
  const stages = [
    { label:'AI Tickets', val:m.aiTickets, pct:100, bg:'linear-gradient(135deg,#00a8cc,#0076cc)' },
    { label:'AI Only', val:m.aiOnly, pct:pct(m.aiOnly,m.aiTickets), bg:'linear-gradient(135deg,#00cc88,#00a86b)' },
    { label:'Human Assisted', val:m.humanAssisted, pct:pct(m.humanAssisted,m.aiTickets), bg:'linear-gradient(135deg,#ffb930,#ff8c00)' },
    { label:'Escalated', val:m.escalated, pct:pct(m.escalated,m.aiTickets), bg:'linear-gradient(135deg,#ff6b6b,#cc0033)' }
  ];

  const container = document.getElementById('funnelVisual');
  container.innerHTML = stages.map((s,i) => `
    ${i>0?'<div class="funnel-arrow"><i class="fa-solid fa-chevron-right"></i></div>':''}
    <div class="funnel-stage" style="background:${s.bg}" title="${s.label}: ${fmt.num(s.val)} (${fmt.pct(s.pct)})">
      <div class="funnel-stage-label">${s.label}</div>
      <div class="funnel-stage-val">${fmt.num(s.val)}</div>
      <div class="funnel-stage-pct">${fmt.pct(s.pct)}</div>
    </div>
  `).join('');
}

// ══════════════════════════════════════════════════
// CHART HELPERS
// ══════════════════════════════════════════════════
function getChartDefaults() {
  const isDark = document.documentElement.getAttribute('data-theme') !== 'light';
  return {
    textColor: isDark ? '#7a95b8' : '#445566',
    gridColor: isDark ? '#1a2d45' : '#d0dce9',
    bgSurface: isDark ? '#101d30' : '#ffffff'
  };
}

function destroyChart(id) {
  if (STATE.charts[id]) { STATE.charts[id].destroy(); delete STATE.charts[id]; }
}

function getDateBuckets(ticketMap) {
  const buckets = new Map();
  ticketMap.forEach(t => {
    if (t.dateBucket) {
      if (!buckets.has(t.dateBucket)) buckets.set(t.dateBucket, []);
      buckets.get(t.dateBucket).push(t);
    }
  });
  // Sort by date
  return new Map([...buckets.entries()].sort((a,b) => a[0].localeCompare(b[0])));
}

// ══════════════════════════════════════════════════
// EFFECTIVENESS CHARTS
// ══════════════════════════════════════════════════
function renderEffectivenessCharts(m) {
  const { textColor, gridColor } = getChartDefaults();
  const buckets = getDateBuckets(STATE.filteredTickets);
  const labels = [...buckets.keys()].map(d => {
    const dt = new Date(d);
    return dt.toLocaleDateString('en-US',{month:'short',day:'numeric'});
  });

  const aiCounts = [...buckets.values()].map(ts => ts.filter(t=>t.isAITicket).length);
  const aiInteractionCounts = [...buckets.values()].map(ts => ts.reduce((s,t)=>s+t.aiInteractionCount,0));
  const fcrRates = [...buckets.values()].map(ts => {
    const ai = ts.filter(t=>t.isAITicket);
    return ai.length ? pct(ai.filter(t=>t.fcrAchieved).length, ai.length) : 0;
  });
  const humanTouchRates = [...buckets.values()].map(ts => {
    const ai = ts.filter(t=>t.isAITicket);
    return ai.length ? pct(ai.filter(t=>t.humanTouchAfterAI).length, ai.length) : 0;
  });
  const escalationRates = [...buckets.values()].map(ts => {
    const ai = ts.filter(t=>t.isAITicket);
    return ai.length ? pct(ai.filter(t=>t.escalated).length, ai.length) : 0;
  });

  const baseLineOpts = {
    responsive: true, maintainAspectRatio: true,
    interaction: { mode:'index', intersect:false },
    plugins: { legend:{labels:{color:textColor,font:{size:11}}}, tooltip:{backgroundColor:'#0d1626',borderColor:'#1a2d45',borderWidth:1} },
    scales: {
      x: { ticks:{color:textColor,font:{size:10},maxRotation:45}, grid:{color:gridColor} },
      y: { ticks:{color:textColor,font:{size:10}}, grid:{color:gridColor}, beginAtZero:true }
    }
  };

  // AI Tickets Trend
  destroyChart('aiTicketsTrend');
  STATE.charts.aiTicketsTrend = new Chart(document.getElementById('aiTicketsTrend'), {
    type:'bar', data:{
      labels, datasets:[{
        label:'AI Tickets', data:aiCounts,
        backgroundColor:'rgba(0,212,255,0.3)', borderColor:'rgba(0,212,255,0.8)', borderWidth:1
      }]
    }, options:{...baseLineOpts}
  });

  // AI Interactions Trend
  destroyChart('aiInteractionsTrend');
  STATE.charts.aiInteractionsTrend = new Chart(document.getElementById('aiInteractionsTrend'), {
    type:'bar', data:{
      labels, datasets:[{
        label:'AI Interactions', data:aiInteractionCounts,
        backgroundColor:'rgba(155,109,255,0.3)', borderColor:'rgba(155,109,255,0.8)', borderWidth:1
      }]
    }, options:{...baseLineOpts}
  });

  // FCR Trend
  destroyChart('fcrTrend');
  STATE.charts.fcrTrend = new Chart(document.getElementById('fcrTrend'), {
    type:'line', data:{
      labels, datasets:[{
        label:'FCR Rate %', data:fcrRates,
        borderColor:'var(--green)', backgroundColor:'rgba(0,229,150,0.1)',
        fill:true, tension:0.4, pointRadius:3
      }]
    }, options:{
      ...baseLineOpts,
      scales:{
        x:{ ticks:{color:textColor,font:{size:10},maxRotation:45}, grid:{color:gridColor} },
        y:{ ticks:{color:textColor,font:{size:10},callback:v=>v+'%'}, grid:{color:gridColor}, beginAtZero:true, max:100 }
      }
    }
  });

  // Human Touch & Escalation
  destroyChart('humanTrend');
  STATE.charts.humanTrend = new Chart(document.getElementById('humanTrend'), {
    type:'line', data:{
      labels, datasets:[
        { label:'Human Touch %', data:humanTouchRates, borderColor:'var(--amber)', backgroundColor:'rgba(255,185,48,0.1)', fill:true, tension:0.4, pointRadius:3 },
        { label:'Escalation %', data:escalationRates, borderColor:'var(--red)', backgroundColor:'rgba(255,77,106,0.05)', fill:false, tension:0.4, pointRadius:3, borderDash:[4,4] }
      ]
    }, options:{
      ...baseLineOpts,
      scales:{
        x:{ ticks:{color:textColor,font:{size:10},maxRotation:45}, grid:{color:gridColor} },
        y:{ ticks:{color:textColor,font:{size:10},callback:v=>v+'%'}, grid:{color:gridColor}, beginAtZero:true }
      }
    }
  });
}

// ══════════════════════════════════════════════════
// COMPLIANCE SECTION
// ══════════════════════════════════════════════════
function renderComplianceSection(m) {
  document.getElementById('gaugeCompPct').textContent = fmt.pct(m.complianceRate);
  document.getElementById('cv-full').textContent = fmt.num(m.compliant);
  document.getElementById('cv-reason').textContent = fmt.num(m.missingReason);
  document.getElementById('cv-sub').textContent = fmt.num(m.missingSubReason);
  document.getElementById('cv-action').textContent = fmt.num(m.missingAction);

  // Gauge
  renderGauge('complianceGauge', m.complianceRate);

  // Pie
  const { textColor } = getChartDefaults();
  destroyChart('compliancePie');
  STATE.charts.compliancePie = new Chart(document.getElementById('compliancePie'), {
    type:'doughnut', data:{
      labels:['Compliant','Missing Reason','Missing Sub Reason','Missing Action'],
      datasets:[{ data:[m.compliant, m.missingReason, m.missingSubReason, m.missingAction],
        backgroundColor:['rgba(0,229,150,0.7)','rgba(255,77,106,0.7)','rgba(255,185,48,0.7)','rgba(155,109,255,0.7)'],
        borderColor:'var(--bg-card)', borderWidth:2
      }]
    },
    options:{
      responsive:true, plugins:{
        legend:{position:'bottom',labels:{color:textColor,font:{size:11},padding:12}},
        tooltip:{backgroundColor:'#0d1626',borderColor:'#1a2d45',borderWidth:1}
      }
    }
  });

  // Bar by date
  const buckets = getDateBuckets(STATE.filteredTickets);
  const labels = [...buckets.keys()].map(d => new Date(d).toLocaleDateString('en-US',{month:'short',day:'numeric'}));
  const { gridColor } = getChartDefaults();

  destroyChart('complianceBar');
  STATE.charts.complianceBar = new Chart(document.getElementById('complianceBar'), {
    type:'bar', data:{
      labels,
      datasets:[
        { label:'Missing Reason', data:[...buckets.values()].map(ts=>ts.filter(t=>t.isAITicket&&t.missingReason).length), backgroundColor:'rgba(255,77,106,0.7)' },
        { label:'Missing Sub Reason', data:[...buckets.values()].map(ts=>ts.filter(t=>t.isAITicket&&t.missingSubReason).length), backgroundColor:'rgba(255,185,48,0.7)' },
        { label:'Missing Action', data:[...buckets.values()].map(ts=>ts.filter(t=>t.isAITicket&&t.missingAction).length), backgroundColor:'rgba(155,109,255,0.7)' }
      ]
    },
    options:{
      responsive:true, plugins:{legend:{labels:{color:textColor,font:{size:11}}}},
      scales:{
        x:{stacked:true, ticks:{color:textColor,font:{size:10},maxRotation:45}, grid:{color:gridColor}},
        y:{stacked:true, ticks:{color:textColor,font:{size:10}}, grid:{color:gridColor}, beginAtZero:true}
      }
    }
  });
}

function renderGauge(canvasId, percentage) {
  const canvas = document.getElementById(canvasId);
  if (!canvas) return;
  const ctx = canvas.getContext('2d');
  const w = canvas.width, h = canvas.height;
  const cx = w/2, cy = h/2;
  const r = Math.min(w,h)/2 - 20;

  ctx.clearRect(0,0,w,h);

  // Background arc
  ctx.beginPath();
  ctx.arc(cx, cy, r, Math.PI*0.75, Math.PI*2.25);
  ctx.strokeStyle = 'rgba(122,149,184,0.15)';
  ctx.lineWidth = 18;
  ctx.lineCap = 'round';
  ctx.stroke();

  // Value arc
  const pctVal = Math.min(100, Math.max(0, percentage));
  const endAngle = Math.PI*0.75 + (pctVal/100)*(Math.PI*1.5);
  const color = pctVal >= 80 ? '#00e596' : pctVal >= 65 ? '#ffb930' : '#ff4d6a';

  ctx.beginPath();
  ctx.arc(cx, cy, r, Math.PI*0.75, endAngle);
  ctx.strokeStyle = color;
  ctx.lineWidth = 18;
  ctx.lineCap = 'round';
  ctx.stroke();
}

// ══════════════════════════════════════════════════
// DEFECT SECTION
// ══════════════════════════════════════════════════
function renderDefectSection(m) {
  const defects = [
    { label:'Duplicate AI Tickets', val:m.duplicateAITickets, type:'error' },
    { label:'Duplicate AI Interactions', val:m.aiTickets > 0 ? m.duplicateAITickets > 0 ? [...STATE.filteredTickets.values()].reduce((s,t)=>s+t.duplicateAICount,0) : 0 : 0, type:'error' },
    { label:'Same Timestamp Defects', val:m.sameTimestampDefects, type:'error' },
    { label:`Short Interval (<${CONFIG.DEFECT_THRESHOLD_SEC}s)`, val:m.shortIntervalDefects, type:'warn' }
  ];

  document.getElementById('defectGrid').innerHTML = defects.map(d => `
    <div class="defect-card ${d.type==='warn'?'warn':''}">
      <div class="defect-label">${d.label}</div>
      <div class="defect-val ${d.type==='warn'?'warn':''}">${fmt.num(d.val)}</div>
    </div>`).join('');

  // Defect table
  const defectTickets = [...STATE.filteredTickets.values()].filter(t => t.hasDefect);

  if (STATE.datatables.defectTable) {
    STATE.datatables.defectTable.destroy();
    document.getElementById('defectTable').innerHTML = '';
  }
  STATE.datatables.defectTable = $('#defectTable').DataTable({
    data: defectTickets,
    pageLength: 10,
    dom: 'Bfrtip',
    buttons:['csv'],
    columns: [
      { title:'Ticket ID', data:'ticketId', render: d => `<span class="ticket-link" onclick="showTimeline('${d}')">${d}</span>` },
      { title:'OGI', data:'ogi' },
      { title:'Date', data:'createdDate', render: d => fmt.date(d) },
      { title:'AI Interactions', data:'aiInteractionCount' },
      { title:'Duplicate AI', data:'hasDuplicateAI', render: d => d ? badge('YES','red') : badge('No','muted') },
      { title:'Same TS', data:'sameTimestampDefects', render: d => d > 0 ? badge(d,'red') : badge(0,'muted') },
      { title:'Short Interval', data:'shortIntervalDefects', render: d => d > 0 ? badge(d,'amber') : badge(0,'muted') },
      { title:'Reason', data:'reason', render: d => d || '<span style="color:var(--text-muted)">—</span>' }
    ]
  });
}

// ══════════════════════════════════════════════════
// REASON ANALYSIS
// ══════════════════════════════════════════════════
function renderReasonAnalysis(m) {
  const aiTickets = m.aiTickets ? m.aiTickets : [...STATE.filteredTickets.values()].filter(t=>t.isAITicket);
  const tickets = typeof aiTickets === 'number' ? [...STATE.filteredTickets.values()].filter(t=>t.isAITicket) : aiTickets;

  const all = [...STATE.filteredTickets.values()];

  // Top reasons AI handled
  const reasonCountAI = countByField(tickets, 'reason');
  const subReasonCount = countByField(tickets, 'subReason');
  const reasonCountEsc = countByField(tickets.filter(t=>t.escalated), 'reason');
  const reasonCountComp = countByField(tickets.filter(t=>!t.compliant), 'reason');

  renderReasonChart('topReasonsAI', 'topReasonsAITable', reasonCountAI, 'var(--cyan)', 'Top Reasons — AI Handled');
  renderReasonChart('topSubReasons', 'topSubReasonsTable', subReasonCount, 'var(--purple)', 'Top Sub Reasons — AI Handled');
  renderReasonChart('topReasonsEsc', 'topReasonsEscTable', reasonCountEsc, 'var(--amber)', 'Top Reasons — Escalated');
  renderReasonChart('topReasonsComp', 'topReasonsCompTable', reasonCountComp, 'var(--red)', 'Top Reasons — Compliance Failures');
}

function countByField(tickets, field) {
  const counts = {};
  tickets.forEach(t => {
    const val = t[field] || 'Unknown';
    counts[val] = (counts[val]||0) + 1;
  });
  return Object.entries(counts).sort((a,b)=>b[1]-a[1]).slice(0,10);
}

function renderReasonChart(chartId, tableId, data, color, label) {
  const { textColor, gridColor } = getChartDefaults();
  const labels = data.map(d=>d[0]);
  const values = data.map(d=>d[1]);

  destroyChart(chartId);
  if (document.getElementById(chartId)) {
    STATE.charts[chartId] = new Chart(document.getElementById(chartId), {
      type:'bar',
      data:{ labels, datasets:[{ label, data:values, backgroundColor:color.replace(')',',0.6)').replace('var(--cyan)','rgba(0,212,255,0.6)').replace('var(--purple)','rgba(155,109,255,0.6)').replace('var(--amber)','rgba(255,185,48,0.6)').replace('var(--red)','rgba(255,77,106,0.6)'), borderRadius:4 }] },
      options:{
        indexAxis:'y', responsive:true,
        plugins:{ legend:{display:false} },
        scales:{
          x:{ ticks:{color:textColor,font:{size:10}}, grid:{color:gridColor} },
          y:{ ticks:{color:textColor,font:{size:10}} }
        }
      }
    });
  }

  const tableEl = document.getElementById(tableId);
  if (!tableEl) return;
  if (STATE.datatables[tableId]) { STATE.datatables[tableId].destroy(); tableEl.innerHTML=''; }

  STATE.datatables[tableId] = $(`#${tableId}`).DataTable({
    data: data.map((d,i) => ({rank:i+1, name:d[0], count:d[1]})),
    pageLength:10, dom:'frtip', searching:true, ordering:true,
    columns:[
      {title:'#', data:'rank', width:'40px'},
      {title:'Reason / Sub Reason', data:'name'},
      {title:'Count', data:'count', render:d=>fmt.num(d)}
    ]
  });
}

// ══════════════════════════════════════════════════
// MASTER TABLE
// ══════════════════════════════════════════════════
function renderMasterTable(m) {
  const tickets = [...STATE.filteredTickets.values()];

  if (STATE.datatables.masterTable) {
    STATE.datatables.masterTable.destroy();
    document.getElementById('masterTable').innerHTML = '';
  }

  STATE.datatables.masterTable = $('#masterTable').DataTable({
    data: tickets,
    pageLength: 25,
    dom: 'Bfrtip',
    buttons: ['csv', 'excel'],
    scrollX: true,
    columns: [
      { title:'OGI', data:'ogi', width:'100px' },
      { title:'Ticket ID', data:'ticketId', render: d => `<a class="ticket-link" onclick="showTimeline('${d}')">${d}</a>`, width:'110px' },
      { title:'Date', data:'createdDate', render: d => fmt.date(d) },
      { title:'Reason', data:'reason', render: d => d || '<span style="color:var(--text-muted)">—</span>' },
      { title:'Sub Reason', data:'subReason', render: d => d || '<span style="color:var(--text-muted)">—</span>' },
      { title:'Action Taken', data:'actionTaken', render: d => d || '<span style="color:var(--text-muted)">—</span>' },
      { title:'AI Ints', data:'aiInteractionCount', width:'60px' },
      { title:'Human Ints', data:'humanInteractionCount', width:'70px' },
      { title:'FCR', data:'fcrAchieved', width:'80px', render: d => d ? badge('PASS','green') : badge('FAIL','red') },
      { title:'Contained', data:'aiContained', width:'90px', render: (d,_,row) => !row.isAITicket ? badge('N/A','muted') : d ? badge('YES','green') : badge('NO','red') },
      { title:'Compliance', data:'compliant', width:'90px', render: (d,_,row) => !row.isAITicket ? badge('N/A','muted') : d ? badge('PASS','green') : badge('FAIL','red') },
      { title:'Defect', data:'hasDefect', width:'70px', render: d => d ? badge('YES','red') : badge('No','muted') },
      { title:'Escalated', data:'escalated', width:'80px', render: (d,_,row) => !row.isAITicket ? badge('N/A','muted') : d ? badge('YES','amber') : badge('No','muted') }
    ]
  });
}

// ══════════════════════════════════════════════════
// TIMELINE VIEW
// ══════════════════════════════════════════════════
function showTimeline(ticketId) {
  const ticket = STATE.filteredTickets.get(ticketId) || STATE.ticketMap.get(ticketId);
  if (!ticket) return;

  document.getElementById('timelineTicketId').textContent = `Ticket: ${ticketId}`;
  document.getElementById('timelineTicketMeta').textContent =
    `OGI: ${ticket.ogi} · ${ticket.interactions.length} interactions · ${ticket.isAITicket?'AI Ticket':'Non-AI'} · FCR: ${ticket.fcrAchieved?'Pass':'Fail'}`;

  const ints = ticket.interactions;
  const tsCount = {};
  ints.forEach(i => { const k = i.parsedDate; if(k) tsCount[k]=(tsCount[k]||0)+1; });

  const html = ints.map((int, idx) => {
    let dotClass = 'internal';
    if (int.type === CONFIG.AI_INTERACTION_TYPE) dotClass = 'ai';
    else if (CONFIG.CUSTOMER_FACING.includes(int.type) && int.type !== CONFIG.AI_INTERACTION_TYPE) dotClass = 'human';

    const isDuplicate = int.type === CONFIG.AI_INTERACTION_TYPE && ticket.hasDuplicateAI && ints.filter(i=>i.type===CONFIG.AI_INTERACTION_TYPE).indexOf(int) > 0;
    const isSameTs = int.parsedDate && tsCount[int.parsedDate] > 1;
    const isEscalation = idx > 0 && CONFIG.CUSTOMER_FACING.includes(int.type) && int.type !== CONFIG.AI_INTERACTION_TYPE && ticket.escalated;

    if (isDuplicate || isSameTs) dotClass = 'duplicate';
    if (isEscalation) dotClass = 'escalation';

    const flags = [];
    if (isDuplicate) flags.push(badge('DUPLICATE AI','red'));
    if (isSameTs) flags.push(badge('SAME TIMESTAMP','red'));
    if (isEscalation) flags.push(badge('ESCALATION','purple'));
    if (CONFIG.INTERNAL_TYPES.includes(int.type)) flags.push(badge('INTERNAL','muted'));
    if (int.type === CONFIG.AI_INTERACTION_TYPE) flags.push(badge('AI','cyan'));

    const missingComp = (!int.reason && ticket.isAITicket) || (!int.subReason && ticket.isAITicket) || (!int.actionTaken && ticket.isAITicket);
    if (missingComp && ticket.isAITicket) {
      const missing = [];
      if (!int.reason) missing.push('Reason');
      if (!int.subReason) missing.push('Sub Reason');
      if (!int.actionTaken) missing.push('Action');
      if (missing.length) flags.push(badge('Missing: '+missing.join(', '),'amber'));
    }

    return `
      <div class="tl-item">
        <div class="tl-dot ${dotClass}"></div>
        <div class="tl-content">
          <div class="tl-time">${fmt.datetime(int.createdDate)}</div>
          <div class="tl-type">${int.type}</div>
          ${int.reason ? `<div style="font-size:11px;color:var(--text-muted)">Reason: ${int.reason}${int.subReason?' · '+int.subReason:''}</div>` : ''}
          <div class="tl-flags">${flags.join('')}</div>
        </div>
      </div>`;
  }).join('');

  document.getElementById('timelineBody').innerHTML = `<div class="timeline-list">${html}</div>`;
  document.getElementById('timelineModal').style.display = 'flex';
}

// ══════════════════════════════════════════════════
// CEO SUMMARY
// ══════════════════════════════════════════════════
function renderCEOSummary(m) {
  const all = [...STATE.filteredTickets.values()];
  const aiTickets = all.filter(t => t.isAITicket);

  // Top escalation drivers
  const escReasons = countByField(aiTickets.filter(t=>t.escalated), 'reason').slice(0,3);

  // Key observations
  const observations = [];
  if (m.containmentRate >= 70) observations.push(`AI is effectively containing <strong>${fmt.pct(m.containmentRate)}</strong> of interactions — customers are not requiring human follow-up.`);
  else observations.push(`AI containment at <strong>${fmt.pct(m.containmentRate)}</strong> — ${fmt.num(m.humanTouch)} tickets required human intervention after AI.`);

  if (m.complianceRate < 90) observations.push(`Compliance rate of <strong>${fmt.pct(m.complianceRate)}</strong> flags ${fmt.num(m.aiTickets - m.compliant)} AI tickets with incomplete reason/action data.`);
  else observations.push(`Compliance is strong at <strong>${fmt.pct(m.complianceRate)}</strong> — data quality is well-maintained.`);

  if (m.duplicateAITickets > 0) observations.push(`<strong>${fmt.num(m.duplicateAITickets)}</strong> tickets have duplicate AI-Agent Call interactions — potential system defect requiring engineering review.`);
  else observations.push('No duplicate AI interaction defects detected in this period.');

  if (m.fcrRate < 70) observations.push(`FCR at <strong>${fmt.pct(m.fcrRate)}</strong> is below target — ${fmt.num(aiTickets.filter(t=>!t.fcrAchieved).length)} AI tickets required repeat contacts.`);
  else observations.push(`FCR rate of <strong>${fmt.pct(m.fcrRate)}</strong> indicates AI is resolving most customer issues in a single interaction.`);

  // Recommendations
  const recommendations = [];
  if (m.humanTouchRate > 30) recommendations.push(`Investigate the top escalation drivers (${escReasons.map(e=>e[0]).join(', ')}) to build AI handling scripts and reduce human escalation rate from ${fmt.pct(m.humanTouchRate)}.`);
  if (m.complianceRate < 95) recommendations.push(`Audit and enforce data entry compliance — ${fmt.num(m.missingReason)} tickets missing Reason, ${fmt.num(m.missingSubReason)} missing Sub Reason, ${fmt.num(m.missingAction)} missing Action Taken.`);
  if (m.duplicateAITickets > 0) recommendations.push(`Engage engineering to investigate ${fmt.num(m.duplicateAITickets)} duplicate AI interaction patterns — could indicate webhook retry issues or system defects.`);
  if (m.containmentRate < 70) recommendations.push('Review AI agent training for low-containment reason categories to improve self-service resolution before escalation.');
  if (recommendations.length < 3) recommendations.push('Continue monitoring weekly FCR and containment trends — establish baseline thresholds for operational alerts.');

  const card = document.getElementById('ceoSummaryCard');
  const topCompFail = m.missingReason >= m.missingSubReason && m.missingReason >= m.missingAction ? 'Missing Reason' : m.missingSubReason >= m.missingAction ? 'Missing Sub Reason' : 'Missing Action Taken';

  card.innerHTML = `
    <div class="ceo-content">
      <div class="ceo-meta-row">
        <div class="ceo-meta-item"><div class="ceo-meta-label">Records Loaded</div><div class="ceo-meta-val">${fmt.num(STATE.normalizedRows.length)}</div></div>
        <div class="ceo-meta-item"><div class="ceo-meta-label">Unique OGIs</div><div class="ceo-meta-val">${fmt.num(new Set([...STATE.filteredTickets.values()].map(t=>t.ogi)).size)}</div></div>
        <div class="ceo-meta-item"><div class="ceo-meta-label">Unique Ticket IDs</div><div class="ceo-meta-val">${fmt.num(STATE.filteredTickets.size)}</div></div>
        <div class="ceo-meta-item"><div class="ceo-meta-label">AI Tickets Created</div><div class="ceo-meta-val">${fmt.num(m.aiTickets)}</div></div>
      </div>
      <div class="ceo-kpi-row">
        <div class="ceo-kpi-item"><div class="ceo-kpi-label">AI Interactions</div><div class="ceo-kpi-val" style="color:var(--cyan)">${fmt.num(m.aiInteractions)}</div></div>
        <div class="ceo-kpi-item"><div class="ceo-kpi-label">FCR Rate</div><div class="ceo-kpi-val" style="color:${m.fcrRate>=75?'var(--green)':m.fcrRate>=60?'var(--amber)':'var(--red)'}">${fmt.pct(m.fcrRate)}</div></div>
        <div class="ceo-kpi-item"><div class="ceo-kpi-label">Containment Rate</div><div class="ceo-kpi-val" style="color:${m.containmentRate>=70?'var(--green)':m.containmentRate>=50?'var(--amber)':'var(--red)'}">${fmt.pct(m.containmentRate)}</div></div>
        <div class="ceo-kpi-item"><div class="ceo-kpi-label">Human Touch Rate</div><div class="ceo-kpi-val" style="color:${m.humanTouchRate<=30?'var(--green)':m.humanTouchRate<=50?'var(--amber)':'var(--red)'}">${fmt.pct(m.humanTouchRate)}</div></div>
        <div class="ceo-kpi-item"><div class="ceo-kpi-label">Compliance Rate</div><div class="ceo-kpi-val" style="color:${m.complianceRate>=90?'var(--green)':m.complianceRate>=75?'var(--amber)':'var(--red)'}">${fmt.pct(m.complianceRate)}</div></div>
      </div>
      <div class="ceo-sections">
        <div class="ceo-col">
          <h4>📊 Top Escalation Drivers</h4>
          ${escReasons.length ? escReasons.map((e,i)=>`<div class="insight-item"><div class="insight-dot" style="background:var(--amber)"></div><div class="insight-text">${i+1}. <strong>${e[0]}</strong> — ${fmt.num(e[1])} escalations</div></div>`).join('') : '<p style="color:var(--text-muted);font-size:13px">No escalations in period</p>'}
          <h4 style="margin-top:1rem">🚨 Top Compliance Failure</h4>
          <div class="insight-item"><div class="insight-dot" style="background:var(--red)"></div><div class="insight-text">${topCompFail} (${fmt.num(Math.max(m.missingReason,m.missingSubReason,m.missingAction))} tickets)</div></div>
        </div>
        <div class="ceo-col">
          <h4>💡 Key Observations</h4>
          ${observations.map(o=>`<div class="insight-item"><div class="insight-dot"></div><div class="insight-text">${o}</div></div>`).join('')}
        </div>
        <div class="ceo-col">
          <h4>🎯 Recommended Actions</h4>
          ${recommendations.slice(0,4).map((r,i)=>`<div class="insight-item"><div class="insight-dot" style="background:var(--purple)"></div><div class="insight-text">${i+1}. ${r}</div></div>`).join('')}
        </div>
      </div>
    </div>`;
}

// ══════════════════════════════════════════════════
// THEME TOGGLE
// ══════════════════════════════════════════════════
function toggleTheme() {
  const html = document.documentElement;
  const isDark = html.getAttribute('data-theme') === 'dark';
  html.setAttribute('data-theme', isDark ? 'light' : 'dark');
  document.getElementById('themeIcon').className = isDark ? 'fa-solid fa-sun' : 'fa-solid fa-moon';
  STATE.currentTheme = isDark ? 'light' : 'dark';
  // Redraw charts if data exists
  if (STATE.validationPassed && STATE.filteredTickets.size) renderDashboard();
}

// ══════════════════════════════════════════════════
// CHART EXPORT
// ══════════════════════════════════════════════════
function setupChartExports() {
  document.querySelectorAll('.btn-chart-export').forEach(btn => {
    btn.addEventListener('click', e => {
      e.stopPropagation();
      const chartId = btn.getAttribute('data-chart');
      const chart = STATE.charts[chartId];
      if (!chart) return;
      const link = document.createElement('a');
      link.download = chartId + '.png';
      link.href = chart.toBase64Image();
      link.click();
      showToast('Chart exported as PNG', 'success');
    });
  });
}

// ══════════════════════════════════════════════════
// PDF EXPORT
// ══════════════════════════════════════════════════
function exportPDF() {
  const { jsPDF } = window.jspdf;
  const doc = new jsPDF({ orientation:'portrait', unit:'mm', format:'a4' });

  const m = computeMetrics(STATE.filteredTickets);
  const now = new Date().toLocaleDateString('en-GB');

  doc.setFillColor(8,14,26);
  doc.rect(0,0,210,40,'F');
  doc.setTextColor(0,212,255);
  doc.setFontSize(18);
  doc.setFont('helvetica','bold');
  doc.text('Way-Decagon AI Effectiveness & Quality', 15, 18);
  doc.setFontSize(10);
  doc.setTextColor(122,149,184);
  doc.text(`Executive Dashboard Export · ${now}`, 15, 28);
  doc.setTextColor(0,229,150);
  doc.text(`Records: ${fmt.num(STATE.normalizedRows.length)} · AI Tickets: ${fmt.num(m.aiTickets)}`, 15, 36);

  let y = 55;
  doc.setTextColor(30,30,30);
  doc.setFontSize(13); doc.setFont('helvetica','bold');
  doc.text('Executive KPI Summary', 15, y); y+=10;

  const rows = [
    ['AI Tickets Created', fmt.num(m.aiTickets)],
    ['AI Agent Interactions', fmt.num(m.aiInteractions)],
    ['AI FCR Rate', fmt.pct(m.fcrRate)],
    ['AI Containment Rate', fmt.pct(m.containmentRate)],
    ['Human Touch Rate', fmt.pct(m.humanTouchRate)],
    ['Escalated to Human', fmt.num(m.escalated)],
    ['Compliance Rate', fmt.pct(m.complianceRate)],
    ['Duplicate AI Tickets', fmt.num(m.duplicateAITickets)],
    ['Same Timestamp Defects', fmt.num(m.sameTimestampDefects)]
  ];

  doc.autoTable({
    startY: y, head:[['KPI','Value']],
    body: rows, margin:{left:15,right:15},
    headStyles:{fillColor:[8,14,26],textColor:[0,212,255],fontSize:10},
    bodyStyles:{fontSize:10},
    alternateRowStyles:{fillColor:[240,244,248]}
  });

  doc.save(`way_decagon_dashboard_${now.replace(/\//g,'-')}.pdf`);
  showToast('PDF exported', 'success');
}

// ══════════════════════════════════════════════════
// EXPORT SUMMARY
// ══════════════════════════════════════════════════
function exportSummaryText() {
  if (!STATE.filteredTickets.size) return;
  const m = computeMetrics(STATE.filteredTickets);
  const escReasons = countByField([...STATE.filteredTickets.values()].filter(t=>t.isAITicket&&t.escalated),'reason').slice(0,3);

  const text = `WAY-DECAGON EXECUTIVE SUMMARY
Generated: ${new Date().toLocaleString()}
${'='.repeat(50)}

DATA RECONCILIATION
Records Loaded: ${fmt.num(STATE.normalizedRows.length)}
Unique OGIs: ${fmt.num(new Set([...STATE.filteredTickets.values()].map(t=>t.ogi)).size)}
Unique Ticket IDs: ${fmt.num(STATE.filteredTickets.size)}

AI PERFORMANCE METRICS
AI Tickets Created: ${fmt.num(m.aiTickets)}
AI Agent Interactions: ${fmt.num(m.aiInteractions)}
AI FCR Rate: ${fmt.pct(m.fcrRate)}
AI Containment Rate: ${fmt.pct(m.containmentRate)}
Human Touch Rate: ${fmt.pct(m.humanTouchRate)}
Escalated to Human: ${fmt.num(m.escalated)}

COMPLIANCE
Compliance Rate: ${fmt.pct(m.complianceRate)}
Missing Reason: ${fmt.num(m.missingReason)}
Missing Sub Reason: ${fmt.num(m.missingSubReason)}
Missing Action Taken: ${fmt.num(m.missingAction)}

SYSTEM DEFECTS
Duplicate AI Tickets: ${fmt.num(m.duplicateAITickets)}
Same Timestamp Defects: ${fmt.num(m.sameTimestampDefects)}
Short Interval Defects: ${fmt.num(m.shortIntervalDefects)}

TOP ESCALATION DRIVERS
${escReasons.map((e,i)=>`${i+1}. ${e[0]} — ${fmt.num(e[1])} escalations`).join('\n')}
`;

  const blob = new Blob([text], {type:'text/plain'});
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url; a.download = 'way_decagon_executive_summary.txt';
  a.click(); URL.revokeObjectURL(url);
  showToast('Executive summary exported', 'success');
}

// ══════════════════════════════════════════════════
// REASON TABS
// ══════════════════════════════════════════════════
function setupReasonTabs() {
  document.querySelectorAll('.reason-tab').forEach(tab => {
    tab.addEventListener('click', () => {
      document.querySelectorAll('.reason-tab').forEach(t => t.classList.remove('active'));
      document.querySelectorAll('.reason-panel').forEach(p => p.classList.remove('active'));
      tab.classList.add('active');
      const panel = document.getElementById('rp-' + tab.dataset.rt);
      if (panel) panel.classList.add('active');
    });
  });
}

// ══════════════════════════════════════════════════
// COMPLIANCE DRILL-DOWN
// ══════════════════════════════════════════════════
function setupComplianceDrills() {
  document.querySelectorAll('.clickable-drill').forEach(el => {
    el.addEventListener('click', () => {
      const drill = el.dataset.drill;
      const tickets = [...STATE.filteredTickets.values()].filter(t => {
        if (!t.isAITicket) return false;
        if (drill === 'missingReason') return t.missingReason;
        if (drill === 'missingSubReason') return t.missingSubReason;
        if (drill === 'missingAction') return t.missingAction;
        return false;
      });

      const wrap = document.getElementById('complianceDrillWrap');
      document.getElementById('compDrillTitle').textContent = `${el.previousElementSibling?.textContent} — ${fmt.num(tickets.length)} tickets`;

      if (STATE.datatables.compDrillTable) {
        STATE.datatables.compDrillTable.destroy();
        document.getElementById('compDrillTable').innerHTML = '';
      }

      STATE.datatables.compDrillTable = $('#compDrillTable').DataTable({
        data: tickets, pageLength:10, dom:'Bfrtip', buttons:['csv'],
        columns:[
          {title:'Ticket ID', data:'ticketId', render:d=>`<span class="ticket-link" onclick="showTimeline('${d}')">${d}</span>`},
          {title:'OGI', data:'ogi'},
          {title:'Date', data:'createdDate', render:d=>fmt.date(d)},
          {title:'Reason', data:'reason', render:d=>d||badge('MISSING','red')},
          {title:'Sub Reason', data:'subReason', render:d=>d||badge('MISSING','red')},
          {title:'Action Taken', data:'actionTaken', render:d=>d||badge('MISSING','red')}
        ]
      });

      wrap.style.display='block';
      wrap.scrollIntoView({behavior:'smooth', block:'start'});
    });
  });

  document.getElementById('closeCompDrill').addEventListener('click', () => {
    document.getElementById('complianceDrillWrap').style.display='none';
  });
}

// ══════════════════════════════════════════════════
// NAV SCROLL SPY
// ══════════════════════════════════════════════════
function setupScrollSpy() {
  const sections = document.querySelectorAll('.dash-section');
  const navLinks = document.querySelectorAll('.nav-link');

  const observer = new IntersectionObserver(entries => {
    entries.forEach(entry => {
      if (entry.isIntersecting) {
        const id = entry.target.id.replace('section-','');
        navLinks.forEach(link => {
          link.classList.toggle('active', link.dataset.section === id);
        });
      }
    });
  }, { threshold: 0.3 });

  sections.forEach(s => observer.observe(s));
}

// ══════════════════════════════════════════════════
// DEFECT RECALC
// ══════════════════════════════════════════════════
function setupDefectRecalc() {
  document.getElementById('recalcDefectsBtn').addEventListener('click', () => {
    const threshold = parseInt(document.getElementById('defectThreshold').value) || 60;
    CONFIG.DEFECT_THRESHOLD_SEC = threshold;
    // Re-enrich all tickets
    STATE.ticketMap.forEach(ticket => enrichTicket(ticket));
    STATE.filteredTickets.forEach(ticket => enrichTicket(ticket));
    const m = computeMetrics(STATE.filteredTickets);
    renderDefectSection(m);
    showToast(`Defect threshold updated to ${threshold}s`, 'info');
  });
}

// ══════════════════════════════════════════════════
// INIT
// ══════════════════════════════════════════════════
document.addEventListener('DOMContentLoaded', () => {
  setupUpload();
  setupReasonTabs();
  setupScrollSpy();
  setupDefectRecalc();

  // Theme toggle
  document.getElementById('themeToggle').addEventListener('click', toggleTheme);

  // PDF export
  document.getElementById('exportPdfBtn').addEventListener('click', exportPDF);
  document.getElementById('exportSummaryBtn').addEventListener('click', exportSummaryText);

  // Filters
  document.getElementById('applyFilterBtn').addEventListener('click', applyFilters);
  document.getElementById('clearFilterBtn').addEventListener('click', clearFilters);

  // Timeline modal close
  document.getElementById('closeTimelineModal').addEventListener('click', () => {
    document.getElementById('timelineModal').style.display='none';
  });
  document.getElementById('timelineModal').addEventListener('click', e => {
    if (e.target === document.getElementById('timelineModal'))
      document.getElementById('timelineModal').style.display='none';
  });

  // Drill modal close
  document.getElementById('closeDrillModal').addEventListener('click', () => {
    document.getElementById('drillModal').style.display='none';
  });

  // Chart exports (set up after DOM ready)
  setupChartExports();

  // Setup compliance drills after initial load
  document.getElementById('section-compliance').addEventListener('click', () => {
    setupComplianceDrills();
  });

  // Keyboard ESC closes modals
  document.addEventListener('keydown', e => {
    if (e.key === 'Escape') {
      document.getElementById('timelineModal').style.display='none';
      document.getElementById('drillModal').style.display='none';
    }
  });

  showToast('Way-Decagon Dashboard ready — upload CSV or load sample data', 'info', 5000);
});

// Expose for inline calls
window.showTimeline = showTimeline;
