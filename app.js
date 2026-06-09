/* ===================================================
   WAY-DECAGON DASHBOARD — APP.JS v3.0
   Supports: CSV + XLSX | White Theme
   Real column mapping for CS_All_Tickets export
=================================================== */
'use strict';

const CONFIG = {
  AI_INTERACTION_TYPE: 'AI-Agent Call',
  INTERNAL_TYPES: ['TL Review','Manager Review','QC Audit','Select','User Reviews','BBB Reviews','App Feedback','Escalation Handled by TL','Escalation handled by Escalation Team','Escalation handled by Manager','Escalation handled by Ops Team'],
  CUSTOMER_FACING: ['AI-Agent Call','Call','Email','Chat','SMS'],
  DEFECT_THRESHOLD_SEC: 60,
  VERSION: '3.0.0'
};

// Map actual CSV/XLSX column names → canonical names used in logic
const COL = {
  TICKET_ID:    ['Ticket ID','ticket_id','ticketid','TicketID'],
  OGI:          ['OGI','ogi','Order Group ID'],
  INTERACTION:  ['Interaction','Interaction Type','interaction_type','channel'],
  INT_DATE:     ['Interaction date','Interaction Date','Created Date','created_date','Date'],
  INT_ID:       ['Interaction ID','interaction_id','InteractionID'],
  REASON:       ['TKT_IssueReason','Reason','reason','Issue Reason'],
  SUB_REASON:   ['Sub Reason','sub_reason','SubReason'],
  ACTION:       ['Action','Action Taken','action_taken','Resolution'],
  STATUS:       ['Status','status'],
  AGENT:        ['Agent Name','agent_name','Assignee'],
  VERTICAL:     ['Vertical','vertical'],
  SUB_VERTICAL: ['SubVertical','Sub Vertical']
};

function findCol(headers, candidates) {
  for (const c of candidates) {
    const found = headers.find(h => h && h.trim().toLowerCase() === c.toLowerCase());
    if (found) return found;
  }
  return null;
}

const STATE = {
  rawRows: [], ticketMap: new Map(), filteredTickets: new Map(),
  charts: {}, datatables: {}, validationPassed: false,
  colMap: {}
};

// ── UTILS ──
const fmt = {
  num: n => n == null ? '—' : Number(n).toLocaleString(),
  pct: n => n == null ? '—' : Number(n).toFixed(1) + '%',
  date: d => { if (!d) return '—'; const dt = new Date(d); return isNaN(dt) ? String(d) : dt.toLocaleDateString('en-US',{month:'short',day:'numeric',year:'numeric'}); },
  datetime: d => { if (!d) return '—'; const dt = new Date(d); return isNaN(dt) ? String(d) : dt.toLocaleString('en-US',{month:'short',day:'numeric',year:'numeric',hour:'2-digit',minute:'2-digit',second:'2-digit'}); }
};

function showToast(msg, type='info', duration=4000) {
  const tc = document.getElementById('toastContainer');
  const t = document.createElement('div');
  const icons = {success:'fa-circle-check',error:'fa-circle-xmark',info:'fa-circle-info'};
  t.className = `toast toast-${type}`;
  t.innerHTML = `<i class="fa-solid ${icons[type]||icons.info}"></i>${msg}`;
  tc.appendChild(t);
  setTimeout(() => { t.style.animation='fadeOut 0.3s ease forwards'; setTimeout(()=>t.remove(),300); }, duration);
}

function unlockSection(id) { document.getElementById(id)?.classList.add('unlocked'); }
function pct(num, den) { return den > 0 ? (num/den*100) : 0; }
function badge(text, color='muted') { return `<span class="badge badge-${color}">${text}</span>`; }

function parseDate(d) {
  if (!d) return null;
  const dt = new Date(d);
  return isNaN(dt) ? null : dt.getTime();
}

// ── COLUMN MAPPING ──
function buildColMap(headers) {
  const map = {};
  map.ticketId   = findCol(headers, COL.TICKET_ID);
  map.ogi        = findCol(headers, COL.OGI);
  map.interaction= findCol(headers, COL.INTERACTION);
  map.intDate    = findCol(headers, COL.INT_DATE);
  map.intId      = findCol(headers, COL.INT_ID);
  map.reason     = findCol(headers, COL.REASON);
  map.subReason  = findCol(headers, COL.SUB_REASON);
  map.action     = findCol(headers, COL.ACTION);
  map.status     = findCol(headers, COL.STATUS);
  map.agent      = findCol(headers, COL.AGENT);
  map.vertical   = findCol(headers, COL.VERTICAL);
  map.subVertical= findCol(headers, COL.SUB_VERTICAL);
  return map;
}

function getVal(row, colName) {
  if (!colName) return '';
  return String(row[colName] || '').trim();
}

// ── XLSX READER ──
async function readXLSX(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = e => {
      try {
        const data = new Uint8Array(e.target.result);
        const workbook = XLSX.read(data, { type: 'array', cellDates: true });
        const sheetName = workbook.SheetNames[0];
        const sheet = workbook.Sheets[sheetName];
        const rows = XLSX.utils.sheet_to_json(sheet, { defval: '', raw: false });
        resolve(rows);
      } catch(err) { reject(err); }
    };
    reader.onerror = reject;
    reader.readAsArrayBuffer(file);
  });
}

// ── VALIDATION ──
function validateData(rows) {
  if (!rows.length) return { passed: false, checks: {}, counts: {} };
  const headers = Object.keys(rows[0]);
  const cm = buildColMap(headers);
  STATE.colMap = cm;

  const checks = {
    cols:       !!cm.ticketId && !!cm.interaction,
    ticket:     rows.some(r => getVal(r, cm.ticketId)),
    ogi:        rows.some(r => getVal(r, cm.ogi)),
    types:      rows.some(r => getVal(r, cm.interaction)),
    compliance: !!cm.reason || !!cm.subReason || !!cm.action
  };

  const isAI = r => getVal(r, cm.interaction) === CONFIG.AI_INTERACTION_TYPE;
  const isInternal = r => CONFIG.INTERNAL_TYPES.includes(getVal(r, cm.interaction));
  const isHumanCustomer = r => {
    const t = getVal(r, cm.interaction);
    return CONFIG.CUSTOMER_FACING.includes(t) && t !== CONFIG.AI_INTERACTION_TYPE;
  };

  const counts = {
    records: rows.length,
    uniqueOGI: new Set(rows.map(r => getVal(r, cm.ogi)).filter(Boolean)).size,
    uniqueTickets: new Set(rows.map(r => getVal(r, cm.ticketId)).filter(Boolean)).size,
    totalInteractions: rows.length,
    aiInteractions: rows.filter(isAI).length,
    humanInteractions: rows.filter(isHumanCustomer).length,
    internalInteractions: rows.filter(isInternal).length
  };

  return { passed: checks.cols && checks.ticket && checks.types, checks, counts };
}

// ── TICKET MAP ──
function buildTicketMap(rows) {
  const cm = STATE.colMap;
  const map = new Map();

  rows.forEach(row => {
    const ticketId = getVal(row, cm.ticketId);
    if (!ticketId) return;

    if (!map.has(ticketId)) {
      map.set(ticketId, {
        ticketId,
        ogi:        getVal(row, cm.ogi) || 'UNKNOWN',
        createdDate:getVal(row, cm.intDate),
        reason:     getVal(row, cm.reason),
        subReason:  getVal(row, cm.subReason),
        actionTaken:getVal(row, cm.action),
        vertical:   getVal(row, cm.vertical),
        subVertical:getVal(row, cm.subVertical),
        status:     getVal(row, cm.status),
        interactions: []
      });
    }

    const ticket = map.get(ticketId);
    const intType = getVal(row, cm.interaction);
    const intDate = getVal(row, cm.intDate);

    ticket.interactions.push({
      interactionId: getVal(row, cm.intId),
      type:          intType,
      createdDate:   intDate,
      parsedDate:    parseDate(intDate),
      reason:        getVal(row, cm.reason),
      subReason:     getVal(row, cm.subReason),
      actionTaken:   getVal(row, cm.action),
      agent:         getVal(row, cm.agent)
    });

    // Keep ticket-level fields from any non-empty row
    if (!ticket.reason && getVal(row, cm.reason)) ticket.reason = getVal(row, cm.reason);
    if (!ticket.subReason && getVal(row, cm.subReason)) ticket.subReason = getVal(row, cm.subReason);
    if (!ticket.actionTaken && getVal(row, cm.action)) ticket.actionTaken = getVal(row, cm.action);
    if (ticket.ogi === 'UNKNOWN' && getVal(row, cm.ogi)) ticket.ogi = getVal(row, cm.ogi);
  });

  map.forEach(ticket => {
    ticket.interactions.sort((a,b) => (a.parsedDate||0) - (b.parsedDate||0));
    if (ticket.interactions.length && ticket.interactions[0].parsedDate)
      ticket.createdDate = ticket.interactions[0].createdDate;
    enrichTicket(ticket);
  });

  return map;
}

function enrichTicket(ticket) {
  const ints = ticket.interactions;
  const customerFacing = ints.filter(i => CONFIG.CUSTOMER_FACING.includes(i.type));
  const aiInts = ints.filter(i => i.type === CONFIG.AI_INTERACTION_TYPE);
  const humanInts = customerFacing.filter(i => i.type !== CONFIG.AI_INTERACTION_TYPE);

  ticket.isAITicket = aiInts.length > 0;
  ticket.aiInteractionCount = aiInts.length;
  ticket.humanInteractionCount = humanInts.length;
  ticket.internalInteractionCount = ints.filter(i => CONFIG.INTERNAL_TYPES.includes(i.type)).length;
  ticket.customerFacingCount = customerFacing.length;
  ticket.fcrAchieved = customerFacing.length <= 1;

  if (ticket.isAITicket) {
    const firstAI = aiInts[0];
    const humanAfterAI = humanInts.filter(i => (i.parsedDate||0) > (firstAI.parsedDate||0));
    ticket.aiContained = humanAfterAI.length === 0;
    ticket.humanTouchAfterAI = humanAfterAI.length > 0;
    ticket.escalated = humanAfterAI.some(i => ['Call'].includes(i.type));
  } else {
    ticket.aiContained = ticket.humanTouchAfterAI = ticket.escalated = false;
  }

  if (ticket.isAITicket) {
    ticket.missingReason    = !ticket.reason;
    ticket.missingSubReason = !ticket.subReason;
    ticket.missingAction    = !ticket.actionTaken;
    ticket.compliant = !ticket.missingReason && !ticket.missingSubReason && !ticket.missingAction;
  } else {
    ticket.missingReason = ticket.missingSubReason = ticket.missingAction = false;
    ticket.compliant = true;
  }

  ticket.hasDuplicateAI = aiInts.length > 1;
  ticket.duplicateAICount = Math.max(0, aiInts.length - 1);

  const tsSet = new Set();
  let sameTs = 0;
  aiInts.forEach(i => { if (!i.parsedDate) return; if (tsSet.has(i.parsedDate)) sameTs++; else tsSet.add(i.parsedDate); });
  ticket.sameTimestampDefects = sameTs;

  let shortInterval = 0;
  for (let i = 1; i < aiInts.length; i++) {
    const p = aiInts[i-1].parsedDate||0, c = aiInts[i].parsedDate||0;
    if (c > 0 && p > 0 && (c - p) < CONFIG.DEFECT_THRESHOLD_SEC * 1000) shortInterval++;
  }
  ticket.shortIntervalDefects = shortInterval;
  ticket.hasDefect = ticket.hasDuplicateAI || sameTs > 0 || shortInterval > 0;

  if (ticket.createdDate) {
    const d = new Date(ticket.createdDate);
    if (!isNaN(d)) {
      ticket.dateBucket = d.toISOString().slice(0,10);
    }
  }
}

function computeMetrics(ticketMap) {
  const tickets = [...ticketMap.values()];
  const aiTickets = tickets.filter(t => t.isAITicket);
  return {
    totalTickets: tickets.length,
    aiTickets: aiTickets.length,
    aiInteractions: tickets.reduce((s,t)=>s+t.aiInteractionCount,0),
    fcrAchieved: aiTickets.filter(t=>t.fcrAchieved).length,
    fcrRate: pct(aiTickets.filter(t=>t.fcrAchieved).length, aiTickets.length),
    aiContained: aiTickets.filter(t=>t.aiContained).length,
    containmentRate: pct(aiTickets.filter(t=>t.aiContained).length, aiTickets.length),
    humanTouch: aiTickets.filter(t=>t.humanTouchAfterAI).length,
    humanTouchRate: pct(aiTickets.filter(t=>t.humanTouchAfterAI).length, aiTickets.length),
    escalated: aiTickets.filter(t=>t.escalated).length,
    escalationRate: pct(aiTickets.filter(t=>t.escalated).length, aiTickets.length),
    compliant: aiTickets.filter(t=>t.compliant).length,
    complianceRate: pct(aiTickets.filter(t=>t.compliant).length, aiTickets.length),
    missingReason: aiTickets.filter(t=>t.missingReason).length,
    missingSubReason: aiTickets.filter(t=>t.missingSubReason).length,
    missingAction: aiTickets.filter(t=>t.missingAction).length,
    duplicateAITickets: aiTickets.filter(t=>t.hasDuplicateAI).length,
    sameTimestampDefects: aiTickets.reduce((s,t)=>s+t.sameTimestampDefects,0),
    shortIntervalDefects: aiTickets.reduce((s,t)=>s+t.shortIntervalDefects,0),
    aiOnly: aiTickets.filter(t=>t.aiContained && !t.humanTouchAfterAI).length,
    humanAssisted: aiTickets.filter(t=>t.humanTouchAfterAI && !t.escalated).length,
    tickets, aiTickets
  };
}

// ── SAMPLE DATA ──
function generateSampleData(n=500) {
  const subReasons = ['Lot Address Enquiry','Check-out Assistance','Shuttle boarding details','Payment Failed','Booking Modification','QR Code Problem','Lot Full','Refund Request','App Technical Issue','Access Issue'];
  const actions = ['Details provided','Issued Refund','Modified Booking','Reset QR Code','Transferred to Lot','Call Transferred to Supervisor','Processed Cancellation','Opened','Closed'];
  const rows = [];
  for (let i=0; i<n; i++) {
    const tid = 1000000 + i;
    const ogi = `OGI${50000000 + Math.floor(i/2)}`;
    const d = new Date('2026-05-01T00:00:00Z');
    d.setDate(d.getDate() + Math.floor(Math.random()*38));
    d.setHours(8 + Math.floor(Math.random()*14), Math.floor(Math.random()*60));
    const sub = subReasons[Math.floor(Math.random()*subReasons.length)];
    const action = Math.random()>0.1 ? actions[Math.floor(Math.random()*actions.length)] : '';
    const reason = Math.random()>0.08 ? (Math.random()>0.5?'Non Escalated':'Not Escalated') : '';
    rows.push({'Ticket ID':tid,'OGI':ogi,'Interaction':'AI-Agent Call','Interaction date':d.toISOString(),'Interaction ID':2000000+i*3,'TKT_IssueReason':reason,'Sub Reason':Math.random()>0.1?sub:'','Action':action,'Status':'Closed','Vertical':'Parking','SubVertical':'Airport Parking','Agent Name':'AI Agent'});
    if (Math.random()<0.08) { const d2=new Date(d); d2.setSeconds(d2.getSeconds()+Math.floor(Math.random()*90)); rows.push({'Ticket ID':tid,'OGI':ogi,'Interaction':'AI-Agent Call','Interaction date':d2.toISOString(),'Interaction ID':2000000+i*3+10,'TKT_IssueReason':reason,'Sub Reason':sub,'Action':action,'Status':'Closed','Vertical':'Parking','SubVertical':'Airport Parking','Agent Name':'AI Agent'}); }
    if (Math.random()<0.3) { const d3=new Date(d); d3.setMinutes(d3.getMinutes()+8); rows.push({'Ticket ID':tid,'OGI':ogi,'Interaction':'TL Review','Interaction date':d3.toISOString(),'Interaction ID':2000000+i*3+11,'TKT_IssueReason':reason,'Sub Reason':sub,'Action':action,'Status':'Closed','Vertical':'Parking','SubVertical':'Airport Parking','Agent Name':'Team Lead'}); }
    if (Math.random()<0.25) { const d4=new Date(d); d4.setMinutes(d4.getMinutes()+20+Math.floor(Math.random()*60)); rows.push({'Ticket ID':tid,'OGI':ogi,'Interaction':Math.random()>0.5?'Call':'Email','Interaction date':d4.toISOString(),'Interaction ID':2000000+i*3+12,'TKT_IssueReason':reason,'Sub Reason':sub,'Action':'Details provided','Status':'Closed','Vertical':'Parking','SubVertical':'Airport Parking','Agent Name':'Arya J S'}); }
  }
  return rows;
}

// ── FILE UPLOAD ──
function setupUpload() {
  const dropZone = document.getElementById('dropZone');
  const fileInput = document.getElementById('fileInput');

  dropZone.addEventListener('click', () => fileInput.click());
  dropZone.addEventListener('dragover', e => { e.preventDefault(); dropZone.classList.add('drag-over'); });
  dropZone.addEventListener('dragleave', () => dropZone.classList.remove('drag-over'));
  dropZone.addEventListener('drop', e => { e.preventDefault(); dropZone.classList.remove('drag-over'); if (e.dataTransfer.files[0]) processFile(e.dataTransfer.files[0]); });
  fileInput.addEventListener('change', e => { if (e.target.files[0]) processFile(e.target.files[0]); });
  document.getElementById('loadSampleBtn').addEventListener('click', () => {
    showToast('Generating 500 sample records…','info');
    setTimeout(() => processRows(generateSampleData(500), 'sample_data.xlsx'), 100);
  });
}

function processFile(file) {
  const name = file.name.toLowerCase();
  const isXLSX = name.endsWith('.xlsx') || name.endsWith('.xls') || name.endsWith('.xlsm');
  const isCSV = name.endsWith('.csv');
  if (!isXLSX && !isCSV) { showToast('Please upload a CSV or XLSX file','error'); return; }

  showUploadProgress();
  showToast(`Reading ${isXLSX ? 'Excel' : 'CSV'} file…`, 'info');

  if (isXLSX) {
    readXLSX(file).then(rows => processRows(rows, file.name)).catch(err => { showToast('Failed to read Excel file: ' + err.message, 'error'); completeProgress(); });
  } else {
    Papa.parse(file, {
      header: true, skipEmptyLines: true, dynamicTyping: false,
      complete: result => processRows(result.data, file.name),
      error: err => { showToast('CSV parse error: ' + err.message, 'error'); completeProgress(); }
    });
  }
}

function showUploadProgress() {
  const pg = document.getElementById('uploadProgress');
  const bar = document.getElementById('progressBar');
  pg.style.display = 'block';
  let p = 0;
  const iv = setInterval(() => { p = Math.min(p + Math.random()*12, 88); bar.style.width = p + '%'; if (p >= 88) clearInterval(iv); }, 150);
  STATE._piv = iv; STATE._pbar = bar;
}

function completeProgress() {
  clearInterval(STATE._piv);
  if (STATE._pbar) STATE._pbar.style.width = '100%';
  setTimeout(() => { document.getElementById('uploadProgress').style.display = 'none'; }, 600);
}

function processRows(rows, filename) {
  STATE.rawRows = rows;
  document.getElementById('fileStatus').style.display = 'block';
  document.getElementById('statFilename').textContent = filename;
  document.getElementById('statLoaded').textContent = new Date().toLocaleString('en-GB',{day:'2-digit',month:'short',year:'numeric',hour:'2-digit',minute:'2-digit'});
  document.getElementById('statRecords').textContent = rows.length.toLocaleString();
  document.getElementById('statStatus').innerHTML = '<span style="color:#059669">✓ Processed Successfully</span>';

  completeProgress();

  const val = validateData(rows);
  renderValidation(val);

  if (!val.passed) { showToast('Validation failed — required columns missing','error'); return; }

  STATE.ticketMap = buildTicketMap(rows);
  STATE.filteredTickets = new Map(STATE.ticketMap);

  const m = computeMetrics(STATE.ticketMap);
  document.getElementById('heroMeta').innerHTML = `
    <span class="meta-pill"><i class="fa-solid fa-database"></i> ${fmt.num(rows.length)} records</span>
    <span class="meta-pill"><i class="fa-solid fa-ticket"></i> ${fmt.num(m.totalTickets)} tickets</span>
    <span class="meta-pill"><i class="fa-solid fa-robot"></i> ${fmt.num(m.aiTickets)} AI tickets</span>
    <span class="meta-pill" style="color:#86efac"><i class="fa-solid fa-circle-check"></i> Live</span>
  `;

  populateFilters();
  renderDashboard();
  showToast(`Loaded ${fmt.num(rows.length)} records — ${fmt.num(m.aiTickets)} AI tickets`, 'success');
}

// ── VALIDATION RENDER ──
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
    'chk-cols':{'ok':checks.cols,'label':'Required columns present'},
    'chk-ticket':{'ok':checks.ticket,'label':'Ticket IDs present'},
    'chk-ogi':{'ok':checks.ogi,'label':'OGI identifiers present'},
    'chk-types':{'ok':checks.types,'label':'Interaction types present'},
    'chk-compliance':{'ok':checks.compliance,'label':'Compliance fields present'}
  };
  Object.entries(checkMap).forEach(([id,{ok,label}]) => {
    const el = document.getElementById(id);
    el.className = `val-check-item ${ok?'pass':'fail'}`;
    el.innerHTML = `<i class="fa-solid fa-${ok?'circle-check':'circle-xmark'}"></i> ${label}`;
  });

  const badge = document.getElementById('validationBadge');
  badge.className = `validation-badge ${passed?'pass':'fail'}`;
  badge.innerHTML = `<i class="fa-solid fa-${passed?'shield-check':'shield-xmark'}"></i> ${passed?'✓ Validation Passed':'✗ Validation Failed'}`;
  STATE.validationPassed = passed;
  unlockSection('section-validation');
}

// ── FILTERS ──
function populateFilters() {
  const tickets = [...STATE.ticketMap.values()];
  const reasons = [...new Set(tickets.map(t=>t.subReason||t.reason).filter(Boolean))].sort();
  const rs = document.getElementById('filterReason');
  rs.innerHTML = '<option value="">All Sub Reasons</option>';
  reasons.slice(0,50).forEach(r => { rs.innerHTML += `<option value="${r}">${r}</option>`; });

  const verticals = [...new Set(tickets.map(t=>t.subVertical||t.vertical).filter(Boolean))].sort();
  const ts = document.getElementById('filterType');
  ts.innerHTML = '<option value="">All Verticals</option>';
  verticals.forEach(v => { ts.innerHTML += `<option value="${v}">${v}</option>`; });

  const dates = tickets.map(t=>t.createdDate).filter(Boolean).map(d=>new Date(d)).filter(d=>!isNaN(d));
  if (dates.length) {
    const min = new Date(Math.min(...dates)).toISOString().slice(0,10);
    const max = new Date(Math.max(...dates)).toISOString().slice(0,10);
    document.getElementById('filterDateFrom').value = min;
    document.getElementById('filterDateTo').value = max;
  }
}

function applyFilters() {
  const dateFrom = document.getElementById('filterDateFrom').value;
  const dateTo = document.getElementById('filterDateTo').value;
  const reason = document.getElementById('filterReason').value;
  const vertical = document.getElementById('filterType').value;

  STATE.filteredTickets = new Map();
  STATE.ticketMap.forEach((ticket, id) => {
    if (dateFrom) { const d = new Date(ticket.createdDate); if (!isNaN(d) && d < new Date(dateFrom)) return; }
    if (dateTo) { const d = new Date(ticket.createdDate); if (!isNaN(d) && d > new Date(dateTo+'T23:59:59')) return; }
    if (reason && ticket.subReason !== reason && ticket.reason !== reason) return;
    if (vertical && ticket.subVertical !== vertical && ticket.vertical !== vertical) return;
    STATE.filteredTickets.set(id, ticket);
  });
  renderDashboard();
  showToast(`Filter applied — ${fmt.num(STATE.filteredTickets.size)} tickets`, 'info');
}

function clearFilters() {
  STATE.filteredTickets = new Map(STATE.ticketMap);
  populateFilters();
  renderDashboard();
  showToast('Filters cleared', 'info');
}

// ── RENDER DASHBOARD ──
function renderDashboard() {
  const m = computeMetrics(STATE.filteredTickets);
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

// ── KPIs ──
function renderKPIs(m) {
  const kpis = [
    {label:'AI Tickets Created', val:fmt.num(m.aiTickets), pctVal:null, icon:'fa-robot', color:'cyan', tip:'Unique Ticket IDs with at least one AI-Agent Call', level:'Ticket'},
    {label:'AI Agent Interactions', val:fmt.num(m.aiInteractions), pctVal:null, icon:'fa-comments', color:'purple', tip:'Total AI-Agent Call interaction records', level:'Interaction'},
    {label:'AI FCR Rate', val:fmt.pct(m.fcrRate), pctVal:m.fcrRate, icon:'fa-bullseye', color:'green', tip:'AI tickets resolved in single customer-facing contact', level:'Ticket'},
    {label:'AI Containment Rate', val:fmt.pct(m.containmentRate), pctVal:m.containmentRate, icon:'fa-shield-halved', color:'green', tip:'AI tickets with no human follow-up after AI', level:'Ticket'},
    {label:'Escalated to Human', val:fmt.num(m.escalated), pctVal:m.escalationRate, icon:'fa-person-walking-arrow-right', color:'amber', tip:'AI tickets where customer called back after AI', level:'Ticket'},
    {label:'Human Touch Rate', val:fmt.pct(m.humanTouchRate), pctVal:m.humanTouchRate, icon:'fa-hand-holding', color:'amber', tip:'AI tickets receiving any human customer contact after AI', level:'Ticket'},
    {label:'Compliance Rate', val:fmt.pct(m.complianceRate), pctVal:m.complianceRate, icon:'fa-clipboard-check', color:'green', tip:'AI tickets with Reason + Sub Reason + Action all filled', level:'Ticket'},
    {label:'Duplicate AI Tickets', val:fmt.num(m.duplicateAITickets), pctVal:null, icon:'fa-copy', color:'red', tip:'Tickets with more than one AI-Agent Call interaction', level:'Ticket'},
    {label:'Same Timestamp Defects', val:fmt.num(m.sameTimestampDefects), pctVal:null, icon:'fa-clock', color:'red', tip:'AI interactions with identical timestamps on same ticket', level:'Interaction'},
    {label:'Compliance Failures', val:fmt.num(m.aiTickets - m.compliant), pctVal:null, icon:'fa-triangle-exclamation', color:'red', tip:'AI tickets missing any compliance field', level:'Ticket'}
  ];
  const colorMap = {
    cyan:{a:'var(--cyan)',d:'var(--cyan-dim)'},
    purple:{a:'var(--purple)',d:'var(--purple-dim)'},
    green:{a:'var(--green)',d:'var(--green-dim)'},
    amber:{a:'var(--amber)',d:'var(--amber-dim)'},
    red:{a:'var(--red)',d:'var(--red-dim)'}
  };
  document.getElementById('kpiGrid').innerHTML = kpis.map(k => {
    const c = colorMap[k.color]||colorMap.cyan;
    const lvlClass = k.level==='Interaction'?'interaction':'';
    return `<div class="kpi-card" style="--accent-color:${c.a};--accent-dim:${c.d}">
      <div class="kpi-tooltip" title="${k.tip}"><i class="fa-solid fa-circle-info"></i></div>
      <div class="kpi-icon"><i class="fa-solid ${k.icon}"></i></div>
      <div class="kpi-label">${k.label}</div>
      <div class="kpi-value">${k.val}</div>
      ${k.pctVal!=null?`<div class="kpi-pct">${fmt.pct(k.pctVal)}</div>`:''}
      <div class="kpi-level-badge ${lvlClass}">${k.level} Level</div>
    </div>`;
  }).join('');
}

// ── FUNNEL ──
function renderFunnel(m) {
  const stages = [
    {label:'AI Tickets',val:m.aiTickets,pct:100,bg:'linear-gradient(135deg,#0ea5e9,#0369a1)'},
    {label:'AI Only',val:m.aiOnly,pct:pct(m.aiOnly,m.aiTickets),bg:'linear-gradient(135deg,#10b981,#059669)'},
    {label:'Human Assisted',val:m.humanAssisted,pct:pct(m.humanAssisted,m.aiTickets),bg:'linear-gradient(135deg,#f59e0b,#d97706)'},
    {label:'Escalated',val:m.escalated,pct:pct(m.escalated,m.aiTickets),bg:'linear-gradient(135deg,#ef4444,#dc2626)'}
  ];
  document.getElementById('funnelVisual').innerHTML = stages.map((s,i) =>
    `${i>0?'<div class="funnel-arrow"><i class="fa-solid fa-chevron-right"></i></div>':''}
    <div class="funnel-stage" style="background:${s.bg}">
      <div class="funnel-stage-label">${s.label}</div>
      <div class="funnel-stage-val">${fmt.num(s.val)}</div>
      <div class="funnel-stage-pct">${fmt.pct(s.pct)}</div>
    </div>`).join('');
}

// ── CHART HELPERS ──
function getChartColors() {
  return { text:'#64748b', grid:'#e2e8f0' };
}
function destroyChart(id) { if (STATE.charts[id]) { STATE.charts[id].destroy(); delete STATE.charts[id]; } }

function getDateBuckets(ticketMap) {
  const b = new Map();
  ticketMap.forEach(t => { if (t.dateBucket) { if (!b.has(t.dateBucket)) b.set(t.dateBucket,[]); b.get(t.dateBucket).push(t); } });
  return new Map([...b.entries()].sort((a,b)=>a[0].localeCompare(b[0])));
}

// ── EFFECTIVENESS CHARTS ──
function renderEffectivenessCharts(m) {
  const {text,grid} = getChartColors();
  const buckets = getDateBuckets(STATE.filteredTickets);
  const labels = [...buckets.keys()].map(d => new Date(d).toLocaleDateString('en-US',{month:'short',day:'numeric'}));
  const aiCounts = [...buckets.values()].map(ts=>ts.filter(t=>t.isAITicket).length);
  const aiIntCounts = [...buckets.values()].map(ts=>ts.reduce((s,t)=>s+t.aiInteractionCount,0));
  const fcrRates = [...buckets.values()].map(ts=>{const ai=ts.filter(t=>t.isAITicket);return ai.length?pct(ai.filter(t=>t.fcrAchieved).length,ai.length):0;});
  const htRates = [...buckets.values()].map(ts=>{const ai=ts.filter(t=>t.isAITicket);return ai.length?pct(ai.filter(t=>t.humanTouchAfterAI).length,ai.length):0;});
  const escRates = [...buckets.values()].map(ts=>{const ai=ts.filter(t=>t.isAITicket);return ai.length?pct(ai.filter(t=>t.escalated).length,ai.length):0;});

  const base = {
    responsive:true, maintainAspectRatio:true,
    interaction:{mode:'index',intersect:false},
    plugins:{legend:{labels:{color:text,font:{size:11}}}, tooltip:{backgroundColor:'#fff',titleColor:'#0f172a',bodyColor:'#475569',borderColor:'#e2e8f0',borderWidth:1}},
    scales:{x:{ticks:{color:text,font:{size:10},maxRotation:45},grid:{color:grid}},y:{ticks:{color:text,font:{size:10}},grid:{color:grid},beginAtZero:true}}
  };

  destroyChart('aiTicketsTrend');
  STATE.charts.aiTicketsTrend = new Chart(document.getElementById('aiTicketsTrend'), {type:'bar', data:{labels, datasets:[{label:'AI Tickets',data:aiCounts,backgroundColor:'rgba(2,132,199,0.6)',borderRadius:4}]}, options:{...base}});

  destroyChart('aiInteractionsTrend');
  STATE.charts.aiInteractionsTrend = new Chart(document.getElementById('aiInteractionsTrend'), {type:'bar', data:{labels, datasets:[{label:'AI Interactions',data:aiIntCounts,backgroundColor:'rgba(124,58,237,0.6)',borderRadius:4}]}, options:{...base}});

  destroyChart('fcrTrend');
  STATE.charts.fcrTrend = new Chart(document.getElementById('fcrTrend'), {type:'line', data:{labels, datasets:[{label:'FCR Rate %',data:fcrRates,borderColor:'#059669',backgroundColor:'rgba(5,150,105,0.1)',fill:true,tension:0.4,pointRadius:3,pointBackgroundColor:'#059669'}]}, options:{...base, scales:{x:{ticks:{color:text,font:{size:10},maxRotation:45},grid:{color:grid}},y:{ticks:{color:text,font:{size:10},callback:v=>v+'%'},grid:{color:grid},beginAtZero:true,max:100}}}});

  destroyChart('humanTrend');
  STATE.charts.humanTrend = new Chart(document.getElementById('humanTrend'), {type:'line', data:{labels, datasets:[{label:'Human Touch %',data:htRates,borderColor:'#d97706',backgroundColor:'rgba(217,119,6,0.08)',fill:true,tension:0.4,pointRadius:3},{label:'Escalation %',data:escRates,borderColor:'#dc2626',fill:false,tension:0.4,pointRadius:3,borderDash:[5,4]}]}, options:{...base, scales:{x:{ticks:{color:text,font:{size:10},maxRotation:45},grid:{color:grid}},y:{ticks:{color:text,font:{size:10},callback:v=>v+'%'},grid:{color:grid},beginAtZero:true}}}});
}

// ── COMPLIANCE ──
function renderComplianceSection(m) {
  document.getElementById('gaugeCompPct').textContent = fmt.pct(m.complianceRate);
  document.getElementById('cv-full').textContent = fmt.num(m.compliant);
  document.getElementById('cv-reason').textContent = fmt.num(m.missingReason);
  document.getElementById('cv-sub').textContent = fmt.num(m.missingSubReason);
  document.getElementById('cv-action').textContent = fmt.num(m.missingAction);
  renderGauge('complianceGauge', m.complianceRate);

  const {text,grid} = getChartColors();
  destroyChart('compliancePie');
  STATE.charts.compliancePie = new Chart(document.getElementById('compliancePie'), {
    type:'doughnut',
    data:{labels:['Compliant','Missing Reason','Missing Sub Reason','Missing Action'],datasets:[{data:[m.compliant,m.missingReason,m.missingSubReason,m.missingAction],backgroundColor:['#10b981','#ef4444','#f59e0b','#8b5cf6'],borderColor:'#fff',borderWidth:2}]},
    options:{responsive:true,plugins:{legend:{position:'bottom',labels:{color:text,font:{size:11},padding:10}},tooltip:{backgroundColor:'#fff',titleColor:'#0f172a',bodyColor:'#475569',borderColor:'#e2e8f0',borderWidth:1}}}
  });

  const buckets = getDateBuckets(STATE.filteredTickets);
  const labels = [...buckets.keys()].map(d => new Date(d).toLocaleDateString('en-US',{month:'short',day:'numeric'}));
  destroyChart('complianceBar');
  STATE.charts.complianceBar = new Chart(document.getElementById('complianceBar'), {
    type:'bar',
    data:{labels,datasets:[{label:'Missing Reason',data:[...buckets.values()].map(ts=>ts.filter(t=>t.isAITicket&&t.missingReason).length),backgroundColor:'rgba(239,68,68,0.7)'},{label:'Missing Sub Reason',data:[...buckets.values()].map(ts=>ts.filter(t=>t.isAITicket&&t.missingSubReason).length),backgroundColor:'rgba(245,158,11,0.7)'},{label:'Missing Action',data:[...buckets.values()].map(ts=>ts.filter(t=>t.isAITicket&&t.missingAction).length),backgroundColor:'rgba(139,92,246,0.7)'}]},
    options:{responsive:true,plugins:{legend:{labels:{color:text,font:{size:11}}}},scales:{x:{stacked:true,ticks:{color:text,font:{size:10},maxRotation:45},grid:{color:grid}},y:{stacked:true,ticks:{color:text,font:{size:10}},grid:{color:grid},beginAtZero:true}}}
  });

  setupComplianceDrills();
}

function renderGauge(id, percentage) {
  const canvas = document.getElementById(id);
  if (!canvas) return;
  const ctx = canvas.getContext('2d'), w=canvas.width, h=canvas.height, cx=w/2, cy=h/2, r=Math.min(w,h)/2-18;
  ctx.clearRect(0,0,w,h);
  ctx.beginPath(); ctx.arc(cx,cy,r,Math.PI*0.75,Math.PI*2.25); ctx.strokeStyle='#e2e8f0'; ctx.lineWidth=16; ctx.lineCap='round'; ctx.stroke();
  const p = Math.min(100,Math.max(0,percentage));
  const color = p>=80?'#10b981':p>=65?'#f59e0b':'#ef4444';
  ctx.beginPath(); ctx.arc(cx,cy,r,Math.PI*0.75,Math.PI*0.75+(p/100)*Math.PI*1.5); ctx.strokeStyle=color; ctx.lineWidth=16; ctx.lineCap='round'; ctx.stroke();
}

// ── DEFECTS ──
function renderDefectSection(m) {
  const totalDupInts = [...STATE.filteredTickets.values()].reduce((s,t)=>s+t.duplicateAICount,0);
  const defects = [
    {label:'Duplicate AI Tickets',val:m.duplicateAITickets,type:'error'},
    {label:'Duplicate AI Interactions',val:totalDupInts,type:'error'},
    {label:'Same Timestamp Defects',val:m.sameTimestampDefects,type:'error'},
    {label:`Short Interval (<${CONFIG.DEFECT_THRESHOLD_SEC}s)`,val:m.shortIntervalDefects,type:'warn'}
  ];
  document.getElementById('defectGrid').innerHTML = defects.map(d=>`<div class="defect-card ${d.type==='warn'?'warn':''}"><div class="defect-label">${d.label}</div><div class="defect-val ${d.type==='warn'?'warn':''}">${fmt.num(d.val)}</div></div>`).join('');

  const defectTickets = [...STATE.filteredTickets.values()].filter(t=>t.hasDefect);
  if (STATE.datatables.defectTable) { STATE.datatables.defectTable.destroy(); document.getElementById('defectTable').innerHTML=''; }
  STATE.datatables.defectTable = $('#defectTable').DataTable({
    data:defectTickets, pageLength:10, dom:'Bfrtip', buttons:['csv'],
    columns:[
      {title:'Ticket ID',data:'ticketId',render:d=>`<span class="ticket-link" onclick="showTimeline('${d}')">${d}</span>`},
      {title:'Vertical',data:'subVertical',render:(d,_,row)=>d||row.vertical||'—'},
      {title:'Date',data:'createdDate',render:d=>fmt.date(d)},
      {title:'AI Ints',data:'aiInteractionCount'},
      {title:'Duplicate',data:'hasDuplicateAI',render:d=>d?badge('YES','red'):badge('No','muted')},
      {title:'Same TS',data:'sameTimestampDefects',render:d=>d>0?badge(d,'red'):badge(0,'muted')},
      {title:'Short Interval',data:'shortIntervalDefects',render:d=>d>0?badge(d,'amber'):badge(0,'muted')}
    ]
  });
}

// ── REASONS ──
function renderReasonAnalysis(m) {
  const aiTickets = [...STATE.filteredTickets.values()].filter(t=>t.isAITicket);
  const countBy = (arr, field) => {
    const c={};
    arr.forEach(t=>{const v=t[field]||'Unknown';c[v]=(c[v]||0)+1;});
    return Object.entries(c).sort((a,b)=>b[1]-a[1]).slice(0,10);
  };
  renderReasonChart('topReasonsAI','topReasonsAITable',countBy(aiTickets,'subReason'),'rgba(2,132,199,0.7)','Top Sub Reasons — AI');
  renderReasonChart('topSubReasons','topSubReasonsTable',countBy(aiTickets,'reason'),'rgba(124,58,237,0.7)','Top Issue Reasons — AI');
  renderReasonChart('topReasonsEsc','topReasonsEscTable',countBy(aiTickets.filter(t=>t.escalated),'subReason'),'rgba(217,119,6,0.7)','Top Sub Reasons — Escalated');
  renderReasonChart('topReasonsComp','topReasonsCompTable',countBy(aiTickets.filter(t=>!t.compliant),'subReason'),'rgba(220,38,38,0.7)','Top Sub Reasons — Compliance Failures');
}

function renderReasonChart(chartId, tableId, data, color, label) {
  const {text,grid} = getChartColors();
  destroyChart(chartId);
  if (document.getElementById(chartId)) {
    STATE.charts[chartId] = new Chart(document.getElementById(chartId), {
      type:'bar',
      data:{labels:data.map(d=>d[0]),datasets:[{label,data:data.map(d=>d[1]),backgroundColor:color,borderRadius:4}]},
      options:{indexAxis:'y',responsive:true,plugins:{legend:{display:false},tooltip:{backgroundColor:'#fff',titleColor:'#0f172a',bodyColor:'#475569',borderColor:'#e2e8f0',borderWidth:1}},scales:{x:{ticks:{color:text,font:{size:10}},grid:{color:grid}},y:{ticks:{color:text,font:{size:10}}}}}
    });
  }
  const el = document.getElementById(tableId);
  if (!el) return;
  if (STATE.datatables[tableId]) { STATE.datatables[tableId].destroy(); el.innerHTML=''; }
  STATE.datatables[tableId] = $(`#${tableId}`).DataTable({
    data:data.map((d,i)=>({rank:i+1,name:d[0],count:d[1]})),
    pageLength:10, dom:'frtip',
    columns:[{title:'#',data:'rank',width:'40px'},{title:'Reason / Sub Reason',data:'name'},{title:'Count',data:'count',render:d=>fmt.num(d)}]
  });
}

// ── MASTER TABLE ──
function renderMasterTable(m) {
  const tickets = [...STATE.filteredTickets.values()];
  if (STATE.datatables.masterTable) { STATE.datatables.masterTable.destroy(); document.getElementById('masterTable').innerHTML=''; }
  STATE.datatables.masterTable = $('#masterTable').DataTable({
    data:tickets, pageLength:25, dom:'Bfrtip', buttons:['csv','excel'], scrollX:true,
    columns:[
      {title:'OGI',data:'ogi',width:'110px'},
      {title:'Ticket ID',data:'ticketId',render:d=>`<a class="ticket-link" onclick="showTimeline('${d}')">${d}</a>`,width:'100px'},
      {title:'Date',data:'createdDate',render:d=>fmt.date(d)},
      {title:'Vertical',data:'subVertical',render:(d,_,row)=>d||row.vertical||'—'},
      {title:'Sub Reason',data:'subReason',render:d=>d||'<span style="color:#94a3b8">—</span>'},
      {title:'Action',data:'actionTaken',render:d=>d||'<span style="color:#94a3b8">—</span>'},
      {title:'AI',data:'aiInteractionCount',width:'50px'},
      {title:'Human',data:'humanInteractionCount',width:'60px'},
      {title:'FCR',data:'fcrAchieved',width:'75px',render:d=>d?badge('PASS','green'):badge('FAIL','red')},
      {title:'Contained',data:'aiContained',width:'85px',render:(d,_,row)=>!row.isAITicket?badge('N/A','muted'):d?badge('YES','green'):badge('NO','red')},
      {title:'Compliance',data:'compliant',width:'90px',render:(d,_,row)=>!row.isAITicket?badge('N/A','muted'):d?badge('PASS','green'):badge('FAIL','red')},
      {title:'Defect',data:'hasDefect',width:'65px',render:d=>d?badge('YES','red'):badge('No','muted')},
      {title:'Escalated',data:'escalated',width:'80px',render:(d,_,row)=>!row.isAITicket?badge('N/A','muted'):d?badge('YES','amber'):badge('No','muted')}
    ]
  });
}

// ── TIMELINE ──
function showTimeline(ticketId) {
  const ticket = STATE.filteredTickets.get(String(ticketId)) || STATE.ticketMap.get(String(ticketId));
  if (!ticket) return;

  document.getElementById('timelineTicketId').textContent = `Ticket: ${ticketId}`;
  document.getElementById('timelineTicketMeta').textContent = `OGI: ${ticket.ogi} · ${ticket.interactions.length} interactions · FCR: ${ticket.fcrAchieved?'Pass':'Fail'} · ${ticket.subVertical||ticket.vertical||''}`;

  const tsCount = {};
  ticket.interactions.forEach(i => { if (i.parsedDate) tsCount[i.parsedDate]=(tsCount[i.parsedDate]||0)+1; });
  const aiInts = ticket.interactions.filter(i=>i.type===CONFIG.AI_INTERACTION_TYPE);

  const html = ticket.interactions.map((int, idx) => {
    let dotClass = CONFIG.INTERNAL_TYPES.includes(int.type)?'internal':int.type===CONFIG.AI_INTERACTION_TYPE?'ai':CONFIG.CUSTOMER_FACING.includes(int.type)?'human':'internal';
    const isDup = int.type===CONFIG.AI_INTERACTION_TYPE && aiInts.indexOf(int)>0;
    const isSameTs = int.parsedDate && tsCount[int.parsedDate]>1;
    const isEsc = idx>0 && int.type==='Call' && ticket.escalated;
    if (isDup||isSameTs) dotClass='duplicate';
    if (isEsc) dotClass='escalation';
    const flags = [];
    if (isDup) flags.push(badge('DUPLICATE AI','red'));
    if (isSameTs) flags.push(badge('SAME TIMESTAMP','red'));
    if (isEsc) flags.push(badge('ESCALATION POINT','purple'));
    if (CONFIG.INTERNAL_TYPES.includes(int.type)) flags.push(badge('INTERNAL','muted'));
    if (int.type===CONFIG.AI_INTERACTION_TYPE) flags.push(badge('AI','cyan'));
    return `<div class="tl-item"><div class="tl-dot ${dotClass}"></div><div class="tl-content"><div class="tl-time">${fmt.datetime(int.createdDate)}</div><div class="tl-type">${int.type}</div>${int.subReason?`<div style="font-size:11px;color:#64748b">${int.subReason}</div>`:''}<div class="tl-flags">${flags.join('')}</div></div></div>`;
  }).join('');

  document.getElementById('timelineBody').innerHTML = `<div class="timeline-list">${html}</div>`;
  document.getElementById('timelineModal').style.display = 'flex';
}

// ── COMPLIANCE DRILLS ──
function setupComplianceDrills() {
  document.querySelectorAll('.clickable-drill').forEach(el => {
    el.onclick = () => {
      const drill = el.dataset.drill;
      const tickets = [...STATE.filteredTickets.values()].filter(t => {
        if (!t.isAITicket) return false;
        if (drill==='missingReason') return t.missingReason;
        if (drill==='missingSubReason') return t.missingSubReason;
        if (drill==='missingAction') return t.missingAction;
        return false;
      });
      const wrap = document.getElementById('complianceDrillWrap');
      document.getElementById('compDrillTitle').textContent = `${el.previousElementSibling?.textContent||'Drill'} — ${fmt.num(tickets.length)} tickets`;
      if (STATE.datatables.compDrillTable) { STATE.datatables.compDrillTable.destroy(); document.getElementById('compDrillTable').innerHTML=''; }
      STATE.datatables.compDrillTable = $('#compDrillTable').DataTable({
        data:tickets, pageLength:10, dom:'Bfrtip', buttons:['csv'],
        columns:[
          {title:'Ticket ID',data:'ticketId',render:d=>`<span class="ticket-link" onclick="showTimeline('${d}')">${d}</span>`},
          {title:'OGI',data:'ogi'},
          {title:'Date',data:'createdDate',render:d=>fmt.date(d)},
          {title:'Sub Reason',data:'subReason',render:d=>d||badge('MISSING','red')},
          {title:'Action Taken',data:'actionTaken',render:d=>d||badge('MISSING','red')}
        ]
      });
      wrap.style.display='block';
      wrap.scrollIntoView({behavior:'smooth',block:'start'});
    };
  });
  document.getElementById('closeCompDrill').onclick = () => { document.getElementById('complianceDrillWrap').style.display='none'; };
}

// ── CEO SUMMARY ──
function renderCEOSummary(m) {
  const aiTickets = [...STATE.filteredTickets.values()].filter(t=>t.isAITicket);
  const countBy = (arr,field) => { const c={}; arr.forEach(t=>{const v=t[field]||'Unknown';c[v]=(c[v]||0)+1;}); return Object.entries(c).sort((a,b)=>b[1]-a[1]); };
  const escReasons = countBy(aiTickets.filter(t=>t.escalated),'subReason').slice(0,3);

  const obs = [];
  if (m.containmentRate>=70) obs.push(`AI is containing <strong>${fmt.pct(m.containmentRate)}</strong> of tickets without human intervention.`);
  else obs.push(`AI containment at <strong>${fmt.pct(m.containmentRate)}</strong> — ${fmt.num(m.humanTouch)} tickets needed human follow-up.`);
  if (m.complianceRate<90) obs.push(`Compliance at <strong>${fmt.pct(m.complianceRate)}</strong> — ${fmt.num(m.aiTickets-m.compliant)} AI tickets have incomplete data fields.`);
  else obs.push(`Compliance strong at <strong>${fmt.pct(m.complianceRate)}</strong>.`);
  if (m.duplicateAITickets>0) obs.push(`<strong>${fmt.num(m.duplicateAITickets)}</strong> duplicate AI interaction tickets detected — potential system defect.`);
  if (m.fcrRate<70) obs.push(`FCR at <strong>${fmt.pct(m.fcrRate)}</strong> — ${fmt.num(aiTickets.filter(t=>!t.fcrAchieved).length)} AI tickets had repeat customer contacts.`);
  else obs.push(`FCR at <strong>${fmt.pct(m.fcrRate)}</strong> — strong first-contact resolution.`);

  const recs = [];
  if (m.humanTouchRate>30) recs.push(`Analyse top escalation drivers (${escReasons.map(e=>e[0]).slice(0,2).join(', ')}) — build AI scripts to reduce ${fmt.pct(m.humanTouchRate)} human touch rate.`);
  if (m.complianceRate<95) recs.push(`Enforce compliance tagging — ${fmt.num(m.missingSubReason)} tickets missing Sub Reason, ${fmt.num(m.missingAction)} missing Action.`);
  if (m.duplicateAITickets>0) recs.push(`Engage Engineering on ${fmt.num(m.duplicateAITickets)} duplicate AI ticket patterns — likely webhook/retry issue.`);
  if (recs.length<3) recs.push('Establish weekly FCR and containment rate baselines for ongoing AI performance monitoring.');

  const topFail = m.missingSubReason>=m.missingReason&&m.missingSubReason>=m.missingAction?`Missing Sub Reason (${fmt.num(m.missingSubReason)})`:m.missingReason>=m.missingAction?`Missing Reason (${fmt.num(m.missingReason)})`:`Missing Action (${fmt.num(m.missingAction)})`;
  const kpiColor = (v,good,mid) => v>=good?'#059669':v>=mid?'#d97706':'#dc2626';

  document.getElementById('ceoSummaryCard').innerHTML = `<div class="ceo-content">
    <div class="ceo-meta-row">
      <div class="ceo-meta-item"><div class="ceo-meta-label">Records Loaded</div><div class="ceo-meta-val">${fmt.num(STATE.rawRows.length)}</div></div>
      <div class="ceo-meta-item"><div class="ceo-meta-label">Unique OGIs</div><div class="ceo-meta-val">${fmt.num(new Set([...STATE.filteredTickets.values()].map(t=>t.ogi)).size)}</div></div>
      <div class="ceo-meta-item"><div class="ceo-meta-label">Unique Ticket IDs</div><div class="ceo-meta-val">${fmt.num(STATE.filteredTickets.size)}</div></div>
      <div class="ceo-meta-item"><div class="ceo-meta-label">AI Tickets</div><div class="ceo-meta-val">${fmt.num(m.aiTickets)}</div></div>
    </div>
    <div class="ceo-kpi-row">
      <div class="ceo-kpi-item"><div class="ceo-kpi-label">AI Interactions</div><div class="ceo-kpi-val" style="color:#7c3aed">${fmt.num(m.aiInteractions)}</div></div>
      <div class="ceo-kpi-item"><div class="ceo-kpi-label">FCR Rate</div><div class="ceo-kpi-val" style="color:${kpiColor(m.fcrRate,75,60)}">${fmt.pct(m.fcrRate)}</div></div>
      <div class="ceo-kpi-item"><div class="ceo-kpi-label">Containment Rate</div><div class="ceo-kpi-val" style="color:${kpiColor(m.containmentRate,70,50)}">${fmt.pct(m.containmentRate)}</div></div>
      <div class="ceo-kpi-item"><div class="ceo-kpi-label">Human Touch Rate</div><div class="ceo-kpi-val" style="color:${m.humanTouchRate<=30?'#059669':m.humanTouchRate<=50?'#d97706':'#dc2626'}">${fmt.pct(m.humanTouchRate)}</div></div>
      <div class="ceo-kpi-item"><div class="ceo-kpi-label">Compliance Rate</div><div class="ceo-kpi-val" style="color:${kpiColor(m.complianceRate,90,75)}">${fmt.pct(m.complianceRate)}</div></div>
    </div>
    <div class="ceo-sections">
      <div class="ceo-col">
        <h4>📊 Top Escalation Drivers</h4>
        ${escReasons.length?escReasons.map((e,i)=>`<div class="insight-item"><div class="insight-dot" style="background:#d97706"></div><div class="insight-text">${i+1}. <strong>${e[0]}</strong> — ${fmt.num(e[1])}</div></div>`).join(''):'<p style="font-size:12px;color:#64748b">No escalations in period</p>'}
        <h4 style="margin-top:1rem">🚨 Top Compliance Failure</h4>
        <div class="insight-item"><div class="insight-dot" style="background:#dc2626"></div><div class="insight-text">${topFail}</div></div>
      </div>
      <div class="ceo-col">
        <h4>💡 Key Observations</h4>
        ${obs.map(o=>`<div class="insight-item"><div class="insight-dot"></div><div class="insight-text">${o}</div></div>`).join('')}
      </div>
      <div class="ceo-col">
        <h4>🎯 Recommended Actions</h4>
        ${recs.map((r,i)=>`<div class="insight-item"><div class="insight-dot" style="background:#7c3aed"></div><div class="insight-text">${i+1}. ${r}</div></div>`).join('')}
      </div>
    </div>
  </div>`;
}

// ── PDF EXPORT ──
function exportPDF() {
  if (!STATE.filteredTickets.size) { showToast('Upload data first','error'); return; }
  const {jsPDF} = window.jspdf;
  const doc = new jsPDF({orientation:'portrait',unit:'mm',format:'a4'});
  const m = computeMetrics(STATE.filteredTickets);
  const now = new Date().toLocaleDateString('en-GB');
  doc.setFillColor(15,23,42); doc.rect(0,0,210,35,'F');
  doc.setTextColor(255,255,255); doc.setFontSize(16); doc.setFont('helvetica','bold');
  doc.text('Way-Decagon AI Effectiveness & Quality Dashboard',15,16);
  doc.setFontSize(9); doc.setTextColor(148,163,184);
  doc.text(`Executive Export · ${now} · Records: ${fmt.num(STATE.rawRows.length)}`,15,26);
  let y=48;
  doc.setTextColor(15,23,42); doc.setFontSize(12); doc.text('Executive KPI Summary',15,y); y+=8;
  doc.autoTable({startY:y,head:[['KPI','Value']],body:[
    ['AI Tickets Created',fmt.num(m.aiTickets)],['AI Agent Interactions',fmt.num(m.aiInteractions)],
    ['AI FCR Rate',fmt.pct(m.fcrRate)],['AI Containment Rate',fmt.pct(m.containmentRate)],
    ['Human Touch Rate',fmt.pct(m.humanTouchRate)],['Escalated to Human',fmt.num(m.escalated)],
    ['Compliance Rate',fmt.pct(m.complianceRate)],['Compliance Failures',fmt.num(m.aiTickets-m.compliant)],
    ['Duplicate AI Tickets',fmt.num(m.duplicateAITickets)]
  ],margin:{left:15,right:15},headStyles:{fillColor:[15,23,42],textColor:[255,255,255],fontSize:9},bodyStyles:{fontSize:9}});
  doc.save(`way_decagon_${now.replace(/\//g,'-')}.pdf`);
  showToast('PDF exported','success');
}

function exportSummary() {
  if (!STATE.filteredTickets.size) return;
  const m = computeMetrics(STATE.filteredTickets);
  const text = `WAY-DECAGON EXECUTIVE SUMMARY\nGenerated: ${new Date().toLocaleString()}\n${'='.repeat(50)}\nAI Tickets: ${fmt.num(m.aiTickets)}\nFCR Rate: ${fmt.pct(m.fcrRate)}\nContainment Rate: ${fmt.pct(m.containmentRate)}\nHuman Touch Rate: ${fmt.pct(m.humanTouchRate)}\nCompliance Rate: ${fmt.pct(m.complianceRate)}\nDuplicate AI Tickets: ${fmt.num(m.duplicateAITickets)}\n`;
  const blob = new Blob([text],{type:'text/plain'});
  const a = document.createElement('a'); a.href=URL.createObjectURL(blob); a.download='way_decagon_summary.txt'; a.click();
  showToast('Summary exported','success');
}

// ── REASON TABS ──
function setupReasonTabs() {
  document.querySelectorAll('.reason-tab').forEach(tab => {
    tab.addEventListener('click', () => {
      document.querySelectorAll('.reason-tab').forEach(t=>t.classList.remove('active'));
      document.querySelectorAll('.reason-panel').forEach(p=>p.classList.remove('active'));
      tab.classList.add('active');
      document.getElementById('rp-'+tab.dataset.rt)?.classList.add('active');
    });
  });
}

// ── SCROLL SPY ──
function setupScrollSpy() {
  const sections = document.querySelectorAll('.dash-section');
  const navLinks = document.querySelectorAll('.nav-link');
  new IntersectionObserver(entries => {
    entries.forEach(e => { if (e.isIntersecting) { const id=e.target.id.replace('section-',''); navLinks.forEach(l=>l.classList.toggle('active',l.dataset.section===id)); } });
  },{threshold:0.3}).observe || sections.forEach(s => new IntersectionObserver(entries => entries.forEach(e=>{if(e.isIntersecting){const id=e.target.id.replace('section-','');navLinks.forEach(l=>l.classList.toggle('active',l.dataset.section===id));}}),{threshold:0.3}).observe(s));
  const obs = new IntersectionObserver(entries => {
    entries.forEach(e => { if (e.isIntersecting) { const id=e.target.id.replace('section-',''); navLinks.forEach(l=>l.classList.toggle('active',l.dataset.section===id)); }});
  },{threshold:0.3});
  sections.forEach(s=>obs.observe(s));
}

// ── INIT ──
document.addEventListener('DOMContentLoaded', () => {
  setupUpload();
  setupReasonTabs();
  setupScrollSpy();

  document.getElementById('exportPdfBtn').addEventListener('click', exportPDF);
  document.getElementById('exportSummaryBtn').addEventListener('click', exportSummary);
  document.getElementById('applyFilterBtn').addEventListener('click', applyFilters);
  document.getElementById('clearFilterBtn').addEventListener('click', clearFilters);
  document.getElementById('closeTimelineModal').addEventListener('click', ()=>{ document.getElementById('timelineModal').style.display='none'; });
  document.getElementById('timelineModal').addEventListener('click', e=>{ if(e.target===document.getElementById('timelineModal')) document.getElementById('timelineModal').style.display='none'; });
  document.getElementById('closeDrillModal').addEventListener('click', ()=>{ document.getElementById('drillModal').style.display='none'; });
  document.getElementById('recalcDefectsBtn').addEventListener('click', () => {
    CONFIG.DEFECT_THRESHOLD_SEC = parseInt(document.getElementById('defectThreshold').value)||60;
    STATE.ticketMap.forEach(enrichTicket);
    STATE.filteredTickets.forEach(enrichTicket);
    renderDefectSection(computeMetrics(STATE.filteredTickets));
    showToast(`Threshold updated to ${CONFIG.DEFECT_THRESHOLD_SEC}s`,'info');
  });
  document.addEventListener('keydown', e=>{ if(e.key==='Escape'){document.getElementById('timelineModal').style.display='none';document.getElementById('drillModal').style.display='none';}});

  document.querySelectorAll('.btn-chart-export').forEach(btn => {
    btn.addEventListener('click', e => {
      e.stopPropagation();
      const chart = STATE.charts[btn.dataset.chart];
      if (!chart) return;
      const a=document.createElement('a'); a.download=btn.dataset.chart+'.png'; a.href=chart.toBase64Image(); a.click();
      showToast('Chart saved as PNG','success');
    });
  });

  showToast('Dashboard ready — upload your CS All Tickets XLSX or CSV','info',5000);
});

window.showTimeline = showTimeline;
