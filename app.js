/* WAY-DECAGON DASHBOARD v3.1 */
'use strict';

const CONFIG = {
  AI_TYPE: 'AI-Agent Call',
  INTERNAL: ['TL Review','Manager Review','QC Audit','Select','User Reviews','BBB Reviews','App Feedback','Escalation Handled by TL','Escalation handled by Escalation Team','Escalation handled by Manager','Escalation handled by Ops Team'],
  CUSTOMER_FACING: ['AI-Agent Call','Call','Email','Chat','SMS'],
  DEFECT_THRESHOLD_SEC: 60,
  MIN_DATE: '2026-05-20'
};

const STATE = {
  rawRows:[], ticketMap:new Map(), filteredTickets:new Map(),
  charts:{}, datatables:{}, colMap:{},
  currentReason:null, currentReasonTab:'handled',
  reasonData:{}
};

// ── UTILS ──
const fmt = {
  num: n => n==null?'—':Number(n).toLocaleString(),
  pct: n => n==null?'—':Number(n).toFixed(1)+'%',
  date: d => { if(!d) return '—'; const dt=new Date(d); return isNaN(dt)?String(d):dt.toLocaleDateString('en-US',{month:'short',day:'numeric',year:'numeric'}); },
  datetime: d => { if(!d) return '—'; const dt=new Date(d); return isNaN(dt)?String(d):dt.toLocaleString('en-US',{month:'short',day:'numeric',year:'numeric',hour:'2-digit',minute:'2-digit',second:'2-digit'}); }
};

function showToast(msg,type='info',dur=4000){
  const tc=document.getElementById('toastContainer');
  const t=document.createElement('div');
  const icons={success:'fa-circle-check',error:'fa-circle-xmark',info:'fa-circle-info'};
  t.className=`toast toast-${type}`;
  t.innerHTML=`<i class="fa-solid ${icons[type]||icons.info}"></i>${msg}`;
  tc.appendChild(t);
  setTimeout(()=>{t.style.animation='fadeOut 0.3s ease forwards';setTimeout(()=>t.remove(),300);},dur);
}

function pct(n,d){return d>0?(n/d*100):0}
function badge(t,c='muted'){return `<span class="badge badge-${c}">${t}</span>`}
function parseDate(d){if(!d)return null;const dt=new Date(d);return isNaN(dt)?null:dt.getTime();}

// ── COLUMN MAP ──
const COL_CANDIDATES = {
  ticketId:['Ticket ID','ticket_id'],
  ogi:['OGI','ogi'],
  interaction:['Interaction','Interaction Type'],
  intDate:['Interaction date','Interaction Date','Created Date'],
  intId:['Interaction ID'],
  reason:['TKT_IssueReason','Reason'],
  subReason:['Sub Reason','sub_reason'],
  action:['Action','Action Taken'],
  status:['Status'],
  agent:['Agent Name'],
  vertical:['Vertical'],
  subVertical:['SubVertical','Sub Vertical'],
  ticketCreatedDate:['Ticket_created_date','Ticket Created Date']
};

function buildColMap(headers){
  const map={};
  Object.entries(COL_CANDIDATES).forEach(([key,candidates])=>{
    map[key]=candidates.find(c=>headers.find(h=>h&&h.trim().toLowerCase()===c.toLowerCase()))||null;
  });
  return map;
}
function getV(row,col){return col?String(row[col]||'').trim():''}

// ── XLSX READER ──
async function readXLSX(file){
  return new Promise((resolve,reject)=>{
    const reader=new FileReader();
    reader.onload=e=>{
      try{
        document.getElementById('progressLabel').textContent='Parsing Excel...';
        const data=new Uint8Array(e.target.result);
        const wb=XLSX.read(data,{type:'array',cellDates:true,cellNF:false,cellStyles:false,cellFormula:false});
        const ws=wb.Sheets[wb.SheetNames[0]];
        setTimeout(()=>{
          try{resolve(XLSX.utils.sheet_to_json(ws,{defval:'',raw:false}));}
          catch(err){reject(err);}
        },10);
      }catch(err){reject(err);}
    };
    reader.onerror=reject;
    reader.readAsArrayBuffer(file);
  });
}

// ── VALIDATION ──
function validateData(rows){
  if(!rows.length)return{passed:false,checks:{},counts:{}};
  const headers=Object.keys(rows[0]);
  const cm=buildColMap(headers);
  STATE.colMap=cm;
  const checks={
    cols:!!cm.ticketId&&!!cm.interaction,
    ticket:rows.some(r=>getV(r,cm.ticketId)),
    ogi:rows.some(r=>getV(r,cm.ogi)),
    types:rows.some(r=>getV(r,cm.interaction)),
    compliance:!!cm.reason||!!cm.subReason||!!cm.action
  };
  const isAI=r=>getV(r,cm.interaction)===CONFIG.AI_TYPE;
  const isInternal=r=>CONFIG.INTERNAL.includes(getV(r,cm.interaction));
  const isHuman=r=>{const t=getV(r,cm.interaction);return CONFIG.CUSTOMER_FACING.includes(t)&&t!==CONFIG.AI_TYPE;};
  // Interaction type breakdown
  const typeBreakdown={};
  rows.forEach(r=>{const t=getV(r,cm.interaction)||'Unknown';typeBreakdown[t]=(typeBreakdown[t]||0)+1;});
  STATE.interactionBreakdown=typeBreakdown;
  return{
    passed:checks.cols&&checks.ticket&&checks.types,checks,
    counts:{
      records:rows.length,
      uniqueOGI:new Set(rows.map(r=>getV(r,cm.ogi)).filter(Boolean)).size,
      uniqueTickets:new Set(rows.map(r=>getV(r,cm.ticketId)).filter(Boolean)).size,
      totalInteractions:rows.length,
      aiInteractions:rows.filter(isAI).length,
      humanInteractions:rows.filter(isHuman).length,
      internalInteractions:rows.filter(isInternal).length
    }
  };
}

// ── BUILD TICKET MAP ──
function buildTicketMap(rows){
  const cm=STATE.colMap;
  const map=new Map();
  rows.forEach(row=>{
    const tid=getV(row,cm.ticketId);
    if(!tid)return;
    if(!map.has(tid))map.set(tid,{
      ticketId:tid,
      ogi:getV(row,cm.ogi)||'UNKNOWN',
      createdDate:getV(row,cm.ticketCreatedDate)||getV(row,cm.intDate),
      reason:getV(row,cm.reason),
      subReason:getV(row,cm.subReason),
      actionTaken:getV(row,cm.action),
      status:getV(row,cm.status),
      vertical:getV(row,cm.vertical),
      subVertical:getV(row,cm.subVertical),
      interactions:[]
    });
    const tk=map.get(tid);
    tk.interactions.push({
      interactionId:getV(row,cm.intId),
      type:getV(row,cm.interaction),
      createdDate:getV(row,cm.intDate),
      parsedDate:parseDate(getV(row,cm.intDate)),
      reason:getV(row,cm.reason),
      subReason:getV(row,cm.subReason),
      actionTaken:getV(row,cm.action),
      agent:getV(row,cm.agent),
      status:getV(row,cm.status)
    });
    // Keep best values at ticket level
    if(!tk.reason&&getV(row,cm.reason))tk.reason=getV(row,cm.reason);
    if(!tk.subReason&&getV(row,cm.subReason))tk.subReason=getV(row,cm.subReason);
    if(!tk.actionTaken&&getV(row,cm.action))tk.actionTaken=getV(row,cm.action);
    if(!tk.status&&getV(row,cm.status))tk.status=getV(row,cm.status);
    if(tk.ogi==='UNKNOWN'&&getV(row,cm.ogi))tk.ogi=getV(row,cm.ogi);
  });
  map.forEach(tk=>{
    tk.interactions.sort((a,b)=>(a.parsedDate||0)-(b.parsedDate||0));
    enrichTicket(tk);
  });
  return map;
}

function enrichTicket(tk){
  const ints=tk.interactions;
  const cf=ints.filter(i=>CONFIG.CUSTOMER_FACING.includes(i.type));
  const aiInts=ints.filter(i=>i.type===CONFIG.AI_TYPE);
  const humanInts=cf.filter(i=>i.type!==CONFIG.AI_TYPE);

  // Decagon ticket = first customer-facing interaction is AI-Agent Call
  const firstAI=aiInts.length>0?aiInts[0]:null;
  const humanBeforeAI=humanInts.some(i=>(i.parsedDate||0)<(firstAI?firstAI.parsedDate||0:Infinity));
  tk.isDecagonTicket=firstAI!==null&&!humanBeforeAI;
  tk.aiInteractionCount=aiInts.length;
  tk.humanInteractionCount=humanInts.length;
  tk.internalInteractionCount=ints.filter(i=>CONFIG.INTERNAL.includes(i.type)).length;
  tk.customerFacingCount=cf.length;

  // FCR = only 1 customer-facing interaction total
  tk.fcrAchieved=cf.length<=1;

  if(tk.isDecagonTicket){
    const humanAfterAI=humanInts.filter(i=>(i.parsedDate||0)>(firstAI.parsedDate||0));
    // Containment = no human after AI
    tk.decagonContained=humanAfterAI.length===0;
    // CS Assisted = any human after AI (call, email, chat)
    tk.csAssisted=humanAfterAI.length>0;
  }else{
    tk.decagonContained=tk.csAssisted=false;
  }

  // Compliance: reason + subReason + action + status=closed
  if(tk.isDecagonTicket){
    tk.missingReason=!tk.reason;
    tk.missingSubReason=!tk.subReason;
    tk.missingAction=!tk.actionTaken;
    tk.statusNotClosed=tk.status.toLowerCase()!=='closed';
    tk.compliant=!tk.missingReason&&!tk.missingSubReason&&!tk.missingAction&&!tk.statusNotClosed;
  }else{
    tk.missingReason=tk.missingSubReason=tk.missingAction=tk.statusNotClosed=false;
    tk.compliant=true;
  }

  // Defect: same ticket + same timestamp on AI interactions
  const tsSet=new Set();
  let sameTs=0;
  aiInts.forEach(i=>{if(!i.parsedDate)return;if(tsSet.has(i.parsedDate))sameTs++;else tsSet.add(i.parsedDate);});
  tk.sameTimestampDefects=sameTs;
  tk.hasDefect=sameTs>0;

  // Short interval: ticket created within threshold of another Decagon ticket (handled at aggregate level)
  tk.shortIntervalFlag=false;

  // Best reason for display (sub reason first, fallback to reason)
  tk.displayReason=(tk.subReason||tk.reason||'').trim();

  // Date bucket — only from MIN_DATE onwards
  if(tk.createdDate){
    const d=new Date(tk.createdDate);
    if(!isNaN(d)){
      tk.dateBucket=d.toISOString().slice(0,10);
      if(tk.dateBucket<CONFIG.MIN_DATE)tk.dateBucket=null;
    }
  }
}

// ── COMPUTE METRICS ──
function computeMetrics(ticketMap){
  const all=[...ticketMap.values()];
  const dec=all.filter(t=>t.isDecagonTicket);
  const decCount=dec.length;
  const fcrCount=dec.filter(t=>t.fcrAchieved).length;
  const containedCount=dec.filter(t=>t.decagonContained).length;
  const csAssistedCount=dec.filter(t=>t.csAssisted).length;
  const compliantCount=dec.filter(t=>t.compliant).length;
  const missingReason=dec.filter(t=>t.missingReason).length;
  const missingSubReason=dec.filter(t=>t.missingSubReason).length;
  const statusNotClosed=dec.filter(t=>t.statusNotClosed).length;
  const defectCount=dec.filter(t=>t.hasDefect).length;
  return{
    totalTickets:all.length,
    decagonTickets:decCount,
    decagonInteractions:all.reduce((s,t)=>s+t.aiInteractionCount,0),
    fcrCount,fcrRate:pct(fcrCount,decCount),
    containedCount,containmentRate:pct(containedCount,decCount),
    csAssistedCount,csAssistedRate:pct(csAssistedCount,decCount),
    compliantCount,complianceRate:pct(compliantCount,decCount),
    missingReason,missingSubReason,statusNotClosed,
    complianceFailures:decCount-compliantCount,
    defectCount,
    sameTimestampDefects:dec.reduce((s,t)=>s+t.sameTimestampDefects,0),
    all,dec
  };
}

// ── SAMPLE DATA ──
function generateSampleData(n=500){
  const subs=['Shuttle boarding details at the airport','Lot Address Enquiry','General Enquiry','Shuttle timings','Check-out Assistance','Payment Failed','Booking Modification','QR Code Problem','Need for Additional Parking Time','Shuttle boarding details - General'];
  const reasons=['Non Escalated','Not Escalated','Escalated'];
  const actions=['Details provided','Issued Refund','Modified Booking','Reset QR Code','Transferred to Lot','Details provided','Opened'];
  const statuses=['Closed','Open','In Progress','Waiting for OPs'];
  const rows=[];
  for(let i=0;i<n;i++){
    const tid=1000000+i;
    const ogi=`OGI${50000000+Math.floor(i/2)}`;
    const d=new Date('2026-05-20T00:00:00Z');
    d.setDate(d.getDate()+Math.floor(Math.random()*19));
    d.setHours(8+Math.floor(Math.random()*14),Math.floor(Math.random()*60));
    const sub=Math.random()>0.12?subs[Math.floor(Math.random()*subs.length)]:'';
    const reason=Math.random()>0.08?reasons[Math.floor(Math.random()*reasons.length)]:'';
    const action=actions[Math.floor(Math.random()*actions.length)];
    const status=Math.random()>0.3?'Closed':statuses[Math.floor(Math.random()*statuses.length)];
    rows.push({'Ticket ID':tid,'OGI':ogi,'Interaction':'AI-Agent Call','Interaction date':d.toISOString(),'Interaction ID':2000000+i*3,'TKT_IssueReason':reason,'Sub Reason':sub,'Action':action,'Status':status,'Vertical':'Parking','SubVertical':'Airport Parking','Agent Name':'Decagon AI','Ticket_created_date':d.toISOString()});
    if(Math.random()<0.25){const d2=new Date(d);d2.setMinutes(d2.getMinutes()+20+Math.floor(Math.random()*60));rows.push({'Ticket ID':tid,'OGI':ogi,'Interaction':Math.random()>0.5?'Call':'Email','Interaction date':d2.toISOString(),'Interaction ID':2000000+i*3+1,'TKT_IssueReason':reason,'Sub Reason':sub,'Action':'Details provided','Status':'Closed','Vertical':'Parking','SubVertical':'Airport Parking','Agent Name':'Arya J S','Ticket_created_date':d.toISOString()});}
    if(Math.random()<0.08){const d3=new Date(d);rows.push({'Ticket ID':tid,'OGI':ogi,'Interaction':'AI-Agent Call','Interaction date':d3.toISOString(),'Interaction ID':2000000+i*3+2,'TKT_IssueReason':reason,'Sub Reason':sub,'Action':action,'Status':status,'Vertical':'Parking','SubVertical':'Airport Parking','Agent Name':'Decagon AI','Ticket_created_date':d.toISOString()});}
    if(Math.random()<0.3){const d4=new Date(d);d4.setMinutes(d4.getMinutes()+5);rows.push({'Ticket ID':tid,'OGI':ogi,'Interaction':'TL Review','Interaction date':d4.toISOString(),'Interaction ID':2000000+i*3+3,'TKT_IssueReason':reason,'Sub Reason':sub,'Action':action,'Status':status,'Vertical':'Parking','SubVertical':'Airport Parking','Agent Name':'Team Lead','Ticket_created_date':d.toISOString()});}
  }
  return rows;
}

// ── FILE UPLOAD ──
function setupUpload(){
  const dz=document.getElementById('dropZone');
  const fi=document.getElementById('fileInput');
  dz.addEventListener('click',()=>fi.click());
  dz.addEventListener('dragover',e=>{e.preventDefault();dz.classList.add('drag-over');});
  dz.addEventListener('dragleave',()=>dz.classList.remove('drag-over'));
  dz.addEventListener('drop',e=>{e.preventDefault();dz.classList.remove('drag-over');if(e.dataTransfer.files[0])processFile(e.dataTransfer.files[0]);});
  fi.addEventListener('change',e=>{if(e.target.files[0])processFile(e.target.files[0]);});
  document.getElementById('loadSampleBtn').addEventListener('click',()=>{showToast('Generating sample data…','info');setTimeout(()=>processRows(generateSampleData(500),'sample_data.xlsx'),100);});
}

function processFile(file){
  const name=file.name.toLowerCase();
  showUploadProgress();
  if(name.endsWith('.xlsx')||name.endsWith('.xls')){
    showToast('Reading Excel file…','info');
    readXLSX(file).then(rows=>processRows(rows,file.name)).catch(err=>{showToast('Excel read error: '+err.message,'error');completeProgress();});
  }else{
    Papa.parse(file,{header:true,skipEmptyLines:true,dynamicTyping:false,complete:r=>processRows(r.data,file.name),error:err=>{showToast('CSV error: '+err.message,'error');completeProgress();}});
  }
}

function showUploadProgress(){
  const pg=document.getElementById('uploadProgress'),bar=document.getElementById('progressBar');
  pg.style.display='block';let p=0;
  const iv=setInterval(()=>{p=Math.min(p+Math.random()*12,88);bar.style.width=p+'%';if(p>=88)clearInterval(iv);},150);
  STATE._piv=iv;STATE._pbar=bar;
}

function completeProgress(){
  clearInterval(STATE._piv);
  if(STATE._pbar)STATE._pbar.style.width='100%';
  setTimeout(()=>{document.getElementById('uploadProgress').style.display='none';},600);
}

function processRows(rows,filename){
  STATE.rawRows=rows;
  document.getElementById('fileStatus').style.display='block';
  document.getElementById('statFilename').textContent=filename;
  document.getElementById('statLoaded').textContent=new Date().toLocaleString('en-GB',{day:'2-digit',month:'short',year:'numeric',hour:'2-digit',minute:'2-digit'});
  document.getElementById('statRecords').textContent=rows.length.toLocaleString();
  document.getElementById('statStatus').innerHTML='<span style="color:#059669">✓ Processed Successfully</span>';
  completeProgress();

  const val=validateData(rows);
  renderValidation(val);
  if(!val.passed){showToast('Validation failed — required columns missing','error');return;}

  STATE.ticketMap=buildTicketMap(rows);
  STATE.filteredTickets=new Map(STATE.ticketMap);

  // Compute short interval defects
  computeShortIntervalDefects(STATE.ticketMap);

  const m=computeMetrics(STATE.ticketMap);
  // Set date range
  const dates=[...STATE.ticketMap.values()].map(t=>t.createdDate).filter(Boolean).map(d=>new Date(d)).filter(d=>!isNaN(d));
  if(dates.length){
    const minD=new Date(Math.min(...dates)).toISOString().slice(0,10);
    const maxD=new Date(Math.max(...dates)).toISOString().slice(0,10);
    document.getElementById('globalDateFrom').value=minD;
    document.getElementById('globalDateTo').value=maxD;
    document.getElementById('dateRangeBar').style.display='flex';
  }
  document.getElementById('dataBadge').style.display='flex';
  document.getElementById('dataBadgeText').textContent=`${fmt.num(m.decagonTickets)} Decagon Tickets`;

  renderDashboard();
  showToast(`Loaded ${fmt.num(rows.length)} records — ${fmt.num(m.decagonTickets)} Decagon tickets`,'success');
}

function computeShortIntervalDefects(ticketMap){
  // Find Decagon tickets created within threshold of each other
  const decTks=[...ticketMap.values()].filter(t=>t.isDecagonTicket&&t.createdDate);
  decTks.sort((a,b)=>new Date(a.createdDate)-new Date(b.createdDate));
  for(let i=1;i<decTks.length;i++){
    const prev=new Date(decTks[i-1].createdDate);
    const curr=new Date(decTks[i].createdDate);
    if(!isNaN(prev)&&!isNaN(curr)&&(curr-prev)<CONFIG.DEFECT_THRESHOLD_SEC*1000){
      decTks[i].shortIntervalFlag=true;
    }
  }
}

// ── RENDER ALL ──
function renderDashboard(){
  const m=computeMetrics(STATE.filteredTickets);
  renderValidationBreakdown();
  renderKPIs(m);
  renderFunnel(m);
  renderEffectivenessCharts(m);
  renderComplianceSection(m);
  renderDefectSection(m);
  renderReasonAnalysis(m);
  renderMasterTable(m);
  renderCEOSummary(m);
}

// ── VALIDATION RENDER ──
function renderValidation(val){
  const{counts,checks,passed}=val;
  document.getElementById('vv-records').textContent=fmt.num(counts.records);
  document.getElementById('vv-ogi').textContent=fmt.num(counts.uniqueOGI);
  document.getElementById('vv-tickets').textContent=fmt.num(counts.uniqueTickets);
  document.getElementById('vv-interactions').textContent=fmt.num(counts.totalInteractions);
  document.getElementById('vv-ai').textContent=fmt.num(counts.aiInteractions);
  document.getElementById('vv-human').textContent=fmt.num(counts.humanInteractions);
  document.getElementById('vv-internal').textContent=fmt.num(counts.internalInteractions);
  const checkMap={'chk-cols':{ok:checks.cols,label:'Required columns present'},'chk-ticket':{ok:checks.ticket,label:'Ticket IDs present'},'chk-ogi':{ok:checks.ogi,label:'OGI identifiers present'},'chk-types':{ok:checks.types,label:'Interaction types present'},'chk-compliance':{ok:checks.compliance,label:'Compliance fields present'}};
  Object.entries(checkMap).forEach(([id,{ok,label}])=>{const el=document.getElementById(id);el.className=`val-check-item ${ok?'pass':'fail'}`;el.innerHTML=`<i class="fa-solid fa-${ok?'circle-check':'circle-xmark'}"></i> ${label}`;});
  const b=document.getElementById('validationBadge');
  b.className=`val-badge ${passed?'pass':'fail'}`;
  b.innerHTML=`<i class="fa-solid fa-${passed?'shield-check':'shield-xmark'}"></i> ${passed?'✓ Validation Passed':'✗ Validation Failed'}`;
}

function renderValidationBreakdown(){
  const bd=STATE.interactionBreakdown||{};
  const wrap=document.getElementById('intBreakdownWrap');
  wrap.style.display='block';
  const total=Object.values(bd).reduce((s,v)=>s+v,0);
  const data=Object.entries(bd).sort((a,b)=>b[1]-a[1]).map((e,i)=>({rank:i+1,type:e[0],count:e[1],pct:(e[1]/total*100).toFixed(1)+'%'}));
  if(STATE.datatables.intBreakdown){STATE.datatables.intBreakdown.destroy();document.getElementById('intBreakdownTable').innerHTML='';}
  STATE.datatables.intBreakdown=$('#intBreakdownTable').DataTable({
    data,pageLength:20,dom:'frtip',ordering:true,
    columns:[{title:'#',data:'rank',width:'40px'},{title:'Interaction Type',data:'type'},{title:'Count',data:'count',render:d=>fmt.num(d)},{title:'% of Total',data:'pct'}]
  });
  // Update label
  document.querySelector('#intBreakdownWrap .block-title').innerHTML=`Interaction Type Breakdown <span class="level-tag">All ${fmt.num(total)} Interactions</span>`;
}

// ── KPI CARDS ──
function renderKPIs(m){
  const colorMap={cyan:{a:'var(--cyan)',d:'var(--cyan-dim)'},purple:{a:'var(--purple)',d:'var(--purple-dim)'},green:{a:'var(--green)',d:'var(--green-dim)'},amber:{a:'var(--amber)',d:'var(--amber-dim)'},red:{a:'var(--red)',d:'var(--red-dim)'}};
  const kpis=[
    {label:'Tickets Generated by Decagon',val:fmt.num(m.decagonTickets),pv:null,icon:'fa-robot',color:'cyan',tip:'Unique tickets where the first customer-facing interaction was Decagon (AI-Agent Call)',lvl:'Ticket'},
    {label:'Decagon Interactions',val:fmt.num(m.decagonInteractions),pv:null,icon:'fa-comments',color:'purple',tip:'Total count of AI-Agent Call interaction records across all tickets',lvl:'Interaction'},
    {label:'Decagon FCR',val:fmt.pct(m.fcrRate),pv:m.fcrRate,icon:'fa-bullseye',color:'green',tip:'Decagon tickets resolved in a single customer-facing interaction — customer did not need to contact again',lvl:'Ticket'},
    {label:'Decagon Containment Rate',val:fmt.pct(m.containmentRate),pv:m.containmentRate,icon:'fa-shield-halved',color:'green',tip:'Decagon tickets where no human CS agent interaction followed after Decagon handled it',lvl:'Ticket'},
    {label:'CS Assisted',val:fmt.num(m.csAssistedCount),pv:m.csAssistedRate,icon:'fa-person-walking-arrow-right',color:'amber',tip:'Decagon tickets where a human CS agent had to handle the ticket after Decagon (Call, Email or Chat)',lvl:'Ticket'},
    {label:'CS Assisted Rate',val:fmt.pct(m.csAssistedRate),pv:null,icon:'fa-hand-holding',color:'amber',tip:'Percentage of Decagon tickets that required CS agent involvement after Decagon interaction',lvl:'Ticket'},
    {label:'Compliance Rate',val:fmt.pct(m.complianceRate),pv:m.complianceRate,icon:'fa-clipboard-check',color:'green',tip:'Decagon tickets with Reason + Sub Reason + Action Taken all filled AND Status = Closed',lvl:'Ticket'},
    {label:'Decagon Duplicate Ticket',val:fmt.num(m.sameTimestampDefects),pv:null,icon:'fa-copy',color:'red',tip:'Tickets with 2 or more AI-Agent Call interactions having the exact same timestamp — system defect',lvl:'Ticket'},
    {label:'Short Interval Tickets',val:fmt.num([...STATE.filteredTickets.values()].filter(t=>t.shortIntervalFlag).length),pv:null,icon:'fa-stopwatch',color:'amber',tip:`Decagon tickets created within ${CONFIG.DEFECT_THRESHOLD_SEC} seconds of another Decagon ticket — possible system retry`,lvl:'Ticket'},
    {label:'Compliance Failures',val:fmt.num(m.complianceFailures),pv:null,icon:'fa-triangle-exclamation',color:'red',tip:'Decagon tickets missing any of: Reason, Sub Reason, or having Status not equal to Closed',lvl:'Ticket'}
  ];
  document.getElementById('kpiGrid').innerHTML=kpis.map(k=>{
    const c=colorMap[k.color]||colorMap.cyan;
    return `<div class="kpi-card" style="--ac:${c.a};--acd:${c.d}">
      <div class="kpi-tip" title="${k.tip}"><i class="fa-solid fa-circle-info"></i></div>
      <div class="kpi-icon"><i class="fa-solid ${k.icon}"></i></div>
      <div class="kpi-label">${k.label}</div>
      <div class="kpi-value">${k.val}</div>
      ${k.pv!=null?`<div class="kpi-pct">${fmt.pct(k.pv)}</div>`:''}
      <div class="kpi-lvl ${k.lvl==='Interaction'?'int':''}">${k.lvl} Level</div>
    </div>`;
  }).join('');
}

// ── FUNNEL ──
function renderFunnel(m){
  const stages=[
    {label:'Decagon Tickets',val:m.decagonTickets,p:100,bg:'linear-gradient(135deg,#0ea5e9,#0369a1)'},
    {label:'Decagon Handled',val:m.containedCount,p:pct(m.containedCount,m.decagonTickets),bg:'linear-gradient(135deg,#10b981,#059669)'},
    {label:'CS Assisted',val:m.csAssistedCount,p:pct(m.csAssistedCount,m.decagonTickets),bg:'linear-gradient(135deg,#f59e0b,#d97706)'}
  ];
  document.getElementById('funnelVisual').innerHTML=stages.map((s,i)=>
    `${i>0?'<div class="funnel-arrow"><i class="fa-solid fa-chevron-right"></i></div>':''}
    <div class="funnel-stage" style="background:${s.bg}">
      <div class="funnel-stage-label">${s.label}</div>
      <div class="funnel-stage-val">${fmt.num(s.val)}</div>
      <div class="funnel-stage-pct">${fmt.pct(s.p)}</div>
    </div>`).join('');
}

// ── CHARTS ──
function dChart(id){if(STATE.charts[id]){STATE.charts[id].destroy();delete STATE.charts[id];}}
function getCC(){return{text:'#64748b',grid:'#e2e8f0'};}

function getDateBuckets(ticketMap){
  const b=new Map();
  ticketMap.forEach(t=>{
    if(t.dateBucket&&t.dateBucket>=CONFIG.MIN_DATE){
      if(!b.has(t.dateBucket))b.set(t.dateBucket,[]);
      b.get(t.dateBucket).push(t);
    }
  });
  return new Map([...b.entries()].sort((a,b)=>a[0].localeCompare(b[0])));
}

function renderEffectivenessCharts(m){
  const{text,grid}=getCC();
  const buckets=getDateBuckets(STATE.filteredTickets);
  const labels=[...buckets.keys()].map(d=>new Date(d).toLocaleDateString('en-US',{month:'short',day:'numeric'}));
  const decCounts=[...buckets.values()].map(ts=>ts.filter(t=>t.isDecagonTicket).length);
  const intCounts=[...buckets.values()].map(ts=>ts.reduce((s,t)=>s+t.aiInteractionCount,0));
  const fcrRates=[...buckets.values()].map(ts=>{const d=ts.filter(t=>t.isDecagonTicket);return d.length?pct(d.filter(t=>t.fcrAchieved).length,d.length):0;});
  const csRates=[...buckets.values()].map(ts=>{const d=ts.filter(t=>t.isDecagonTicket);return d.length?pct(d.filter(t=>t.csAssisted).length,d.length):0;});
  const containRates=[...buckets.values()].map(ts=>{const d=ts.filter(t=>t.isDecagonTicket);return d.length?pct(d.filter(t=>t.decagonContained).length,d.length):0;});

  const base={responsive:true,maintainAspectRatio:true,interaction:{mode:'index',intersect:false},plugins:{legend:{labels:{color:text,font:{size:11}}},tooltip:{backgroundColor:'#fff',titleColor:'#0f172a',bodyColor:'#475569',borderColor:'#e2e8f0',borderWidth:1}},scales:{x:{ticks:{color:text,font:{size:10},maxRotation:45},grid:{color:grid}},y:{ticks:{color:text,font:{size:10}},grid:{color:grid},beginAtZero:true}}};

  dChart('decagonTicketsTrend');
  STATE.charts.decagonTicketsTrend=new Chart(document.getElementById('decagonTicketsTrend'),{type:'bar',data:{labels,datasets:[{label:'Decagon Tickets',data:decCounts,backgroundColor:'rgba(2,132,199,0.6)',borderRadius:4}]},options:{...base}});

  dChart('decagonIntsTrend');
  STATE.charts.decagonIntsTrend=new Chart(document.getElementById('decagonIntsTrend'),{type:'bar',data:{labels,datasets:[{label:'Decagon Interactions',data:intCounts,backgroundColor:'rgba(124,58,237,0.6)',borderRadius:4}]},options:{...base}});

  dChart('fcrTrend');
  STATE.charts.fcrTrend=new Chart(document.getElementById('fcrTrend'),{type:'line',data:{labels,datasets:[{label:'Decagon FCR %',data:fcrRates,borderColor:'#059669',backgroundColor:'rgba(5,150,105,0.1)',fill:true,tension:0.4,pointRadius:3,pointBackgroundColor:'#059669'}]},options:{...base,scales:{x:{ticks:{color:text,font:{size:10},maxRotation:45},grid:{color:grid}},y:{ticks:{color:text,font:{size:10},callback:v=>v+'%'},grid:{color:grid},beginAtZero:true,max:100}}}});

  dChart('csAssistedTrend');
  STATE.charts.csAssistedTrend=new Chart(document.getElementById('csAssistedTrend'),{type:'line',data:{labels,datasets:[{label:'CS Assisted %',data:csRates,borderColor:'#d97706',backgroundColor:'rgba(217,119,6,0.08)',fill:true,tension:0.4,pointRadius:3},{label:'Containment %',data:containRates,borderColor:'#059669',fill:false,tension:0.4,pointRadius:3,borderDash:[5,4]}]},options:{...base,scales:{x:{ticks:{color:text,font:{size:10},maxRotation:45},grid:{color:grid}},y:{ticks:{color:text,font:{size:10},callback:v=>v+'%'},grid:{color:grid},beginAtZero:true}}}});
}

// ── COMPLIANCE ──
function renderComplianceSection(m){
  document.getElementById('gaugeCompPct').textContent=fmt.pct(m.complianceRate);
  document.getElementById('cv-full').textContent=fmt.num(m.compliantCount);
  document.getElementById('cv-reason').textContent=fmt.num(m.missingReason);
  document.getElementById('cv-sub').textContent=fmt.num(m.missingSubReason);
  document.getElementById('cv-status').textContent=fmt.num(m.statusNotClosed);
  renderGauge('complianceGauge',m.complianceRate);
  const{text,grid}=getCC();

  dChart('compliancePie');
  STATE.charts.compliancePie=new Chart(document.getElementById('compliancePie'),{
    type:'doughnut',
    data:{labels:['Compliant','Missing Reason','Missing Sub Reason','Status Not Closed'],datasets:[{data:[m.compliantCount,m.missingReason,m.missingSubReason,m.statusNotClosed],backgroundColor:['#10b981','#ef4444','#f59e0b','#8b5cf6'],borderColor:'#fff',borderWidth:2}]},
    options:{responsive:true,plugins:{legend:{position:'bottom',labels:{color:text,font:{size:11},padding:10}},tooltip:{backgroundColor:'#fff',titleColor:'#0f172a',bodyColor:'#475569',borderColor:'#e2e8f0',borderWidth:1}}}
  });

  const buckets=getDateBuckets(STATE.filteredTickets);
  const labels=[...buckets.keys()].map(d=>new Date(d).toLocaleDateString('en-US',{month:'short',day:'numeric'}));
  dChart('complianceBar');
  STATE.charts.complianceBar=new Chart(document.getElementById('complianceBar'),{
    type:'bar',
    data:{labels,datasets:[
      {label:'Missing Reason',data:[...buckets.values()].map(ts=>ts.filter(t=>t.isDecagonTicket&&t.missingReason).length),backgroundColor:'rgba(239,68,68,0.7)'},
      {label:'Missing Sub Reason',data:[...buckets.values()].map(ts=>ts.filter(t=>t.isDecagonTicket&&t.missingSubReason).length),backgroundColor:'rgba(245,158,11,0.7)'},
      {label:'Status Not Closed',data:[...buckets.values()].map(ts=>ts.filter(t=>t.isDecagonTicket&&t.statusNotClosed).length),backgroundColor:'rgba(139,92,246,0.7)'}
    ]},
    options:{responsive:true,plugins:{legend:{labels:{color:text,font:{size:11}}}},scales:{x:{stacked:true,ticks:{color:text,font:{size:10},maxRotation:45},grid:{color:grid}},y:{stacked:true,ticks:{color:text,font:{size:10}},grid:{color:grid},beginAtZero:true}}}
  });

  // Setup drills
  document.querySelectorAll('.ci-val.drill').forEach(el=>{
    el.onclick=()=>{
      const drill=el.dataset.drill;
      const tickets=[...STATE.filteredTickets.values()].filter(t=>{
        if(!t.isDecagonTicket)return false;
        if(drill==='missingReason')return t.missingReason;
        if(drill==='missingSubReason')return t.missingSubReason;
        if(drill==='statusNotClosed')return t.statusNotClosed;
        return false;
      });
      const wrap=document.getElementById('compDrillWrap');
      document.getElementById('compDrillTitle').textContent=`${el.previousElementSibling?.textContent} — ${fmt.num(tickets.length)} tickets`;
      if(STATE.datatables.compDrill){STATE.datatables.compDrill.destroy();document.getElementById('compDrillTable').innerHTML='';}
      STATE.datatables.compDrill=$('#compDrillTable').DataTable({
        data:tickets,pageLength:10,dom:'Bfrtip',buttons:['csv'],
        columns:[
          {title:'Ticket ID',data:'ticketId',render:d=>`<span class="ticket-link" onclick="showTimeline('${d}')">${d}</span>`},
          {title:'OGI',data:'ogi'},
          {title:'Date',data:'createdDate',render:d=>fmt.date(d)},
          {title:'Sub Reason',data:'subReason',render:d=>d||badge('MISSING','red')},
          {title:'Status',data:'status',render:d=>d||badge('MISSING','red')}
        ]
      });
      wrap.style.display='block';wrap.scrollIntoView({behavior:'smooth',block:'start'});
    };
  });
  document.getElementById('closeCompDrill').onclick=()=>{document.getElementById('compDrillWrap').style.display='none';};
}

function renderGauge(id,percentage){
  const canvas=document.getElementById(id);if(!canvas)return;
  const ctx=canvas.getContext('2d'),w=canvas.width,h=canvas.height,cx=w/2,cy=h/2,r=Math.min(w,h)/2-18;
  ctx.clearRect(0,0,w,h);
  ctx.beginPath();ctx.arc(cx,cy,r,Math.PI*0.75,Math.PI*2.25);ctx.strokeStyle='#e2e8f0';ctx.lineWidth=16;ctx.lineCap='round';ctx.stroke();
  const p=Math.min(100,Math.max(0,percentage));
  ctx.beginPath();ctx.arc(cx,cy,r,Math.PI*0.75,Math.PI*0.75+(p/100)*Math.PI*1.5);
  ctx.strokeStyle=p>=80?'#10b981':p>=65?'#f59e0b':'#ef4444';ctx.lineWidth=16;ctx.lineCap='round';ctx.stroke();
}

// ── DEFECTS ──
function renderDefectSection(m){
  const shortInt=[...STATE.filteredTickets.values()].filter(t=>t.shortIntervalFlag).length;
  const defects=[
    {label:'Decagon Duplicate Ticket',val:m.sameTimestampDefects,type:'error',tip:'Same ticket with 2+ AI interactions at identical timestamps'},
    {label:'Duplicate AI Interactions',val:[...STATE.filteredTickets.values()].reduce((s,t)=>s+Math.max(0,t.aiInteractionCount-1),0),type:'error',tip:'Extra AI-Agent Call interactions beyond the first on each ticket'},
    {label:`Short Interval Tickets (<${CONFIG.DEFECT_THRESHOLD_SEC}s)`,val:shortInt,type:'warn',tip:'Decagon tickets created within threshold seconds of another Decagon ticket'}
  ];
  document.getElementById('defectGrid').innerHTML=defects.map(d=>`<div class="defect-card ${d.type==='warn'?'warn':''}"><div class="defect-label">${d.label}</div><div class="defect-val ${d.type==='warn'?'warn':''}" title="${d.tip}">${fmt.num(d.val)}</div></div>`).join('');

  const defectTks=[...STATE.filteredTickets.values()].filter(t=>t.hasDefect||t.shortIntervalFlag);
  if(STATE.datatables.defectTable){STATE.datatables.defectTable.destroy();document.getElementById('defectTable').innerHTML='';}
  STATE.datatables.defectTable=$('#defectTable').DataTable({
    data:defectTks,pageLength:10,dom:'Bfrtip',buttons:['csv'],
    columns:[
      {title:'Ticket ID',data:'ticketId',render:d=>`<span class="ticket-link" onclick="showTimeline('${d}')">${d}</span>`},
      {title:'Date',data:'createdDate',render:d=>fmt.date(d)},
      {title:'AI Interactions',data:'aiInteractionCount'},
      {title:'Same Timestamp',data:'sameTimestampDefects',render:d=>d>0?badge(d,'red'):badge(0,'muted')},
      {title:'Short Interval',data:'shortIntervalFlag',render:d=>d?badge('YES','amber'):badge('No','muted')},
      {title:'Sub Reason',data:'subReason',render:d=>d||'—'}
    ]
  });
}

// ── REASON ANALYSIS ──
function renderReasonAnalysis(m){
  const all=[...STATE.filteredTickets.values()];
  const decTks=all.filter(t=>t.isDecagonTicket);

  STATE.reasonData={
    handled:decTks.filter(t=>t.decagonContained),
    cs:decTks.filter(t=>t.csAssisted),
    comp:decTks.filter(t=>!t.compliant)
  };

  renderReasonChart(STATE.currentReasonTab||'handled');

  document.querySelectorAll('.reason-tab').forEach(tab=>{
    tab.onclick=()=>{
      document.querySelectorAll('.reason-tab').forEach(t=>t.classList.remove('active'));
      tab.classList.add('active');
      STATE.currentReasonTab=tab.dataset.rt;
      STATE.currentReason=null;
      renderReasonChart(tab.dataset.rt);
      document.getElementById('reasonDetailSide').innerHTML=`<div class="reason-detail-placeholder"><i class="fa-solid fa-hand-pointer"></i><p>Click any bar to see breakdown</p></div>`;
    };
  });
}

function getDisplayReason(tk){return(tk.subReason||tk.reason||'').trim()||null;}

function countByReason(tickets){
  const c={};
  tickets.forEach(t=>{
    const r=getDisplayReason(t);
    if(!r)return; // skip blank
    c[r]=(c[r]||0)+1;
  });
  return Object.entries(c).sort((a,b)=>b[1]-a[1]).slice(0,10);
}

function renderReasonChart(tab){
  const tickets=STATE.reasonData[tab]||[];
  const data=countByReason(tickets);
  const{text,grid}=getCC();
  const colors={handled:'rgba(2,132,199,0.7)',cs:'rgba(217,119,6,0.7)',comp:'rgba(220,38,38,0.7)'};

  dChart('reasonChart');
  STATE.charts.reasonChart=new Chart(document.getElementById('reasonChart'),{
    type:'bar',
    data:{labels:data.map(d=>d[0]),datasets:[{label:'Tickets',data:data.map(d=>d[1]),backgroundColor:colors[tab]||colors.handled,borderRadius:4}]},
    options:{
      indexAxis:'y',responsive:true,
      onClick:(evt,els)=>{
        if(!els.length)return;
        const label=data[els[0].index][0];
        STATE.currentReason=label;
        showReasonDetail(label,tab);
      },
      plugins:{legend:{display:false},tooltip:{backgroundColor:'#fff',titleColor:'#0f172a',bodyColor:'#475569',borderColor:'#e2e8f0',borderWidth:1}},
      scales:{x:{ticks:{color:text,font:{size:10}},grid:{color:grid}},y:{ticks:{color:text,font:{size:10}}}}
    }
  });
}

function showReasonDetail(reason,tab){
  const allDecTks=[...STATE.filteredTickets.values()].filter(t=>t.isDecagonTicket);
  // Get ALL Decagon tickets for this reason (not just filtered by tab)
  const reasonTks=allDecTks.filter(t=>getDisplayReason(t)===reason);
  const total=reasonTks.length;

  // Count by status
  const statusCounts={};
  reasonTks.forEach(t=>{
    const s=t.status||'Unknown';
    statusCounts[s]=(statusCounts[s]||0)+1;
  });

  // Also count Decagon Handled vs CS Handled
  const decHandled=reasonTks.filter(t=>t.decagonContained).length;
  const csHandled=reasonTks.filter(t=>t.csAssisted).length;

  const statusColors={'Closed':'#10b981','Open':'#ef4444','In Progress':'#f59e0b','Waiting for OPs':'#8b5cf6'};

  const allStatuses=[
    {name:'Decagon Handled',count:decHandled,color:'#0ea5e9'},
    {name:'CS Handled',count:csHandled,color:'#d97706'},
    ...Object.entries(statusCounts).sort((a,b)=>b[1]-a[1]).map(([s,c])=>({name:s,count:c,color:statusColors[s]||'#94a3b8'}))
  ];

  const side=document.getElementById('reasonDetailSide');
  side.innerHTML=`
    <div class="reason-detail-title">${reason}</div>
    <div class="reason-detail-total">${fmt.num(total)}<span>Total Decagon Tickets</span></div>
    <div class="status-breakdown">
      ${allStatuses.map(s=>`
        <div class="status-row">
          <span class="status-name" style="color:${s.color};font-weight:600">${s.name}</span>
          <div class="status-bar-wrap"><div class="status-bar-fill" style="width:${total?pct(s.count,total):0}%;background:${s.color}"></div></div>
          <span class="status-count">${fmt.num(s.count)}</span>
          <span style="font-size:11px;color:#94a3b8;width:40px;text-align:right">${fmt.pct(total?pct(s.count,total):0)}</span>
        </div>`).join('')}
    </div>`;
}

// ── MASTER TABLE ──
function renderMasterTable(m){
  const tickets=[...STATE.filteredTickets.values()];
  if(STATE.datatables.masterTable){STATE.datatables.masterTable.destroy();document.getElementById('masterTable').innerHTML='';}
  STATE.datatables.masterTable=$('#masterTable').DataTable({
    data:tickets,pageLength:25,dom:'Bfrtip',buttons:['csv','excel'],scrollX:true,
    columns:[
      {title:'OGI',data:'ogi',width:'110px'},
      {title:'Ticket ID',data:'ticketId',render:d=>`<a class="ticket-link" onclick="showTimeline('${d}')">${d}</a>`,width:'100px'},
      {title:'Date',data:'createdDate',render:d=>fmt.date(d)},
      {title:'Vertical',data:'subVertical',render:(d,_,r)=>d||r.vertical||'—'},
      {title:'Sub Reason',data:'subReason',render:d=>d||'<span style="color:#94a3b8">—</span>'},
      {title:'Action',data:'actionTaken',render:d=>d||'<span style="color:#94a3b8">—</span>'},
      {title:'Status',data:'status',render:d=>{const c=d==='Closed'?'green':d==='Open'?'red':'amber';return badge(d||'—',c);}},
      {title:'AI Ints',data:'aiInteractionCount',width:'55px'},
      {title:'FCR',data:'fcrAchieved',width:'70px',render:d=>d?badge('PASS','green'):badge('FAIL','red')},
      {title:'Contained',data:'decagonContained',width:'85px',render:(d,_,r)=>!r.isDecagonTicket?badge('N/A','muted'):d?badge('YES','green'):badge('NO','red')},
      {title:'Compliance',data:'compliant',width:'90px',render:(d,_,r)=>!r.isDecagonTicket?badge('N/A','muted'):d?badge('PASS','green'):badge('FAIL','red')},
      {title:'CS Assisted',data:'csAssisted',width:'85px',render:(d,_,r)=>!r.isDecagonTicket?badge('N/A','muted'):d?badge('YES','amber'):badge('No','muted')}
    ]
  });
}

// ── TIMELINE ──
function showTimeline(ticketId){
  const tk=STATE.filteredTickets.get(String(ticketId))||STATE.ticketMap.get(String(ticketId));
  if(!tk)return;
  document.getElementById('timelineTicketId').textContent=`Ticket: ${ticketId}`;
  document.getElementById('timelineTicketMeta').textContent=`OGI: ${tk.ogi} · ${tk.interactions.length} interactions · FCR: ${tk.fcrAchieved?'Pass':'Fail'} · ${tk.subVertical||tk.vertical||''}`;
  const tsCount={};
  tk.interactions.forEach(i=>{if(i.parsedDate)tsCount[i.parsedDate]=(tsCount[i.parsedDate]||0)+1;});
  const html=tk.interactions.map((int,idx)=>{
    let dc=CONFIG.INTERNAL.includes(int.type)?'internal':int.type===CONFIG.AI_TYPE?'ai':CONFIG.CUSTOMER_FACING.includes(int.type)?'human':'internal';
    const isSameTs=int.parsedDate&&tsCount[int.parsedDate]>1;
    const isEsc=idx>0&&CONFIG.CUSTOMER_FACING.includes(int.type)&&int.type!==CONFIG.AI_TYPE&&tk.csAssisted;
    if(isSameTs)dc='duplicate';if(isEsc)dc='escalation';
    const flags=[];
    if(isSameTs)flags.push(badge('SAME TIMESTAMP','red'));
    if(isEsc)flags.push(badge('CS ASSISTED','amber'));
    if(CONFIG.INTERNAL.includes(int.type))flags.push(badge('INTERNAL','muted'));
    if(int.type===CONFIG.AI_TYPE)flags.push(badge('DECAGON','cyan'));
    return `<div class="tl-item"><div class="tl-dot ${dc}"></div><div class="tl-content"><div class="tl-time">${fmt.datetime(int.createdDate)}</div><div class="tl-type">${int.type}</div>${int.subReason?`<div style="font-size:11px;color:#64748b">${int.subReason}</div>`:''}<div class="tl-flags">${flags.join('')}</div></div></div>`;
  }).join('');
  document.getElementById('timelineBody').innerHTML=`<div class="timeline-list">${html}</div>`;
  document.getElementById('timelineModal').style.display='flex';
}

// ── CEO SUMMARY ──
function renderCEOSummary(m){
  const allDec=[...STATE.filteredTickets.values()].filter(t=>t.isDecagonTicket);
  const escReasons=countByReason(allDec.filter(t=>t.csAssisted)).slice(0,3);
  const obs=[];
  if(m.containmentRate>=70)obs.push(`Decagon is containing <strong>${fmt.pct(m.containmentRate)}</strong> of tickets — customers not requiring CS agent follow-up.`);
  else obs.push(`Decagon containment at <strong>${fmt.pct(m.containmentRate)}</strong> — ${fmt.num(m.csAssistedCount)} tickets required CS agent involvement.`);
  if(m.complianceRate<90)obs.push(`Compliance at <strong>${fmt.pct(m.complianceRate)}</strong> — ${fmt.num(m.statusNotClosed)} tickets still open/in progress after Decagon interaction.`);
  else obs.push(`Compliance strong at <strong>${fmt.pct(m.complianceRate)}</strong> — data quality well maintained.`);
  if(m.fcrRate<70)obs.push(`FCR at <strong>${fmt.pct(m.fcrRate)}</strong> — ${fmt.num(m.decagonTickets-m.fcrCount)} Decagon tickets had repeat customer contacts.`);
  else obs.push(`FCR at <strong>${fmt.pct(m.fcrRate)}</strong> — strong first-contact resolution by Decagon.`);
  if(m.sameTimestampDefects>0)obs.push(`<strong>${fmt.num(m.sameTimestampDefects)}</strong> duplicate timestamp defect detected — potential system issue to investigate.`);

  const recs=[];
  if(m.csAssistedRate>30)recs.push(`Analyse top CS-assisted drivers (${escReasons.slice(0,2).map(e=>e[0]).join(', ')}) — train Decagon to handle these without human escalation.`);
  if(m.statusNotClosed>100)recs.push(`${fmt.num(m.statusNotClosed)} Decagon tickets still open/in progress — review resolution workflows to improve ticket closure rate.`);
  if(m.missingSubReason>50)recs.push(`${fmt.num(m.missingSubReason)} tickets missing Sub Reason — enforce tagging compliance in Decagon configuration.`);
  if(recs.length<3)recs.push('Establish weekly Decagon FCR and containment baselines for performance monitoring and alerting.');

  const kpiColor=(v,good,mid)=>v>=good?'#059669':v>=mid?'#d97706':'#dc2626';
  const ogi=new Set([...STATE.filteredTickets.values()].map(t=>t.ogi)).size;

  document.getElementById('ceoSummaryCard').innerHTML=`<div class="ceo-content">
    <div class="ceo-meta-row">
      <div class="ceo-meta-item"><div class="ceo-meta-label">Records Loaded</div><div class="ceo-meta-val">${fmt.num(STATE.rawRows.length)}</div></div>
      <div class="ceo-meta-item"><div class="ceo-meta-label">Unique OGIs</div><div class="ceo-meta-val">${fmt.num(ogi)}</div></div>
      <div class="ceo-meta-item"><div class="ceo-meta-label">Unique Ticket IDs</div><div class="ceo-meta-val">${fmt.num(STATE.filteredTickets.size)}</div></div>
      <div class="ceo-meta-item"><div class="ceo-meta-label">Decagon Tickets</div><div class="ceo-meta-val">${fmt.num(m.decagonTickets)}</div></div>
    </div>
    <div class="ceo-kpi-row">
      <div class="ceo-kpi-item"><div class="ceo-kpi-label">Decagon Interactions</div><div class="ceo-kpi-val" style="color:#7c3aed">${fmt.num(m.decagonInteractions)}</div></div>
      <div class="ceo-kpi-item"><div class="ceo-kpi-label">Decagon FCR</div><div class="ceo-kpi-val" style="color:${kpiColor(m.fcrRate,75,60)}">${fmt.pct(m.fcrRate)}</div></div>
      <div class="ceo-kpi-item"><div class="ceo-kpi-label">Containment Rate</div><div class="ceo-kpi-val" style="color:${kpiColor(m.containmentRate,70,50)}">${fmt.pct(m.containmentRate)}</div></div>
      <div class="ceo-kpi-item"><div class="ceo-kpi-label">CS Assisted Rate</div><div class="ceo-kpi-val" style="color:${m.csAssistedRate<=30?'#059669':m.csAssistedRate<=50?'#d97706':'#dc2626'}">${fmt.pct(m.csAssistedRate)}</div></div>
      <div class="ceo-kpi-item"><div class="ceo-kpi-label">Compliance Rate</div><div class="ceo-kpi-val" style="color:${kpiColor(m.complianceRate,90,75)}">${fmt.pct(m.complianceRate)}</div></div>
    </div>
    <div class="ceo-sections">
      <div class="ceo-col">
        <h4>📊 Top CS-Assisted Drivers</h4>
        ${escReasons.length?escReasons.map((e,i)=>`<div class="insight-item"><div class="insight-dot" style="background:#d97706"></div><div class="insight-text">${i+1}. <strong>${e[0]}</strong> — ${fmt.num(e[1])} tickets</div></div>`).join(''):'<p style="font-size:12px;color:#64748b">No CS-assisted tickets in period</p>'}
        <h4 style="margin-top:1rem">🚨 Top Compliance Issue</h4>
        <div class="insight-item"><div class="insight-dot" style="background:#dc2626"></div><div class="insight-text">Status Not Closed: ${fmt.num(m.statusNotClosed)} tickets</div></div>
        <div class="insight-item"><div class="insight-dot" style="background:#f59e0b"></div><div class="insight-text">Missing Sub Reason: ${fmt.num(m.missingSubReason)} tickets</div></div>
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

// ── DATE FILTER ──
function applyDateFilter(){
  const from=document.getElementById('globalDateFrom').value;
  const to=document.getElementById('globalDateTo').value;
  STATE.filteredTickets=new Map();
  STATE.ticketMap.forEach((tk,id)=>{
    if(from){const d=new Date(tk.createdDate);if(!isNaN(d)&&d<new Date(from))return;}
    if(to){const d=new Date(tk.createdDate);if(!isNaN(d)&&d>new Date(to+'T23:59:59'))return;}
    STATE.filteredTickets.set(id,tk);
  });
  renderDashboard();
  showToast(`Date filter applied — ${fmt.num(STATE.filteredTickets.size)} tickets`,'info');
}

function clearDateFilter(){
  STATE.filteredTickets=new Map(STATE.ticketMap);
  const dates=[...STATE.ticketMap.values()].map(t=>t.createdDate).filter(Boolean).map(d=>new Date(d)).filter(d=>!isNaN(d));
  if(dates.length){
    document.getElementById('globalDateFrom').value=new Date(Math.min(...dates)).toISOString().slice(0,10);
    document.getElementById('globalDateTo').value=new Date(Math.max(...dates)).toISOString().slice(0,10);
  }
  renderDashboard();
  showToast('Date filter cleared','info');
}

// ── EXPORT ──
function exportPDF(){
  if(!STATE.filteredTickets.size){showToast('Upload data first','error');return;}
  const{jsPDF}=window.jspdf;
  const doc=new jsPDF({orientation:'portrait',unit:'mm',format:'a4'});
  const m=computeMetrics(STATE.filteredTickets);
  const now=new Date().toLocaleDateString('en-GB');
  doc.setFillColor(15,23,42);doc.rect(0,0,210,35,'F');
  doc.setTextColor(255,255,255);doc.setFontSize(16);doc.setFont('helvetica','bold');
  doc.text('Way-Decagon AI Effectiveness & Quality Dashboard',15,16);
  doc.setFontSize(9);doc.setTextColor(148,163,184);
  doc.text(`Executive Export · ${now} · Records: ${fmt.num(STATE.rawRows.length)}`,15,26);
  let y=48;
  doc.setTextColor(15,23,42);doc.setFontSize(12);doc.text('Executive KPI Summary',15,y);y+=8;
  doc.autoTable({startY:y,head:[['KPI','Value']],body:[
    ['Decagon Tickets',fmt.num(m.decagonTickets)],['Decagon Interactions',fmt.num(m.decagonInteractions)],
    ['Decagon FCR',fmt.pct(m.fcrRate)],['Decagon Containment Rate',fmt.pct(m.containmentRate)],
    ['CS Assisted Rate',fmt.pct(m.csAssistedRate)],['CS Assisted Count',fmt.num(m.csAssistedCount)],
    ['Compliance Rate',fmt.pct(m.complianceRate)],['Compliance Failures',fmt.num(m.complianceFailures)],
    ['Status Not Closed',fmt.num(m.statusNotClosed)],['Duplicate Timestamp Defects',fmt.num(m.sameTimestampDefects)]
  ],margin:{left:15,right:15},headStyles:{fillColor:[15,23,42],textColor:[255,255,255],fontSize:9},bodyStyles:{fontSize:9}});
  doc.save(`way_decagon_${now.replace(/\//g,'-')}.pdf`);
  showToast('PDF exported','success');
}

function exportSummary(){
  if(!STATE.filteredTickets.size)return;
  const m=computeMetrics(STATE.filteredTickets);
  const text=`WAY-DECAGON EXECUTIVE SUMMARY\nGenerated: ${new Date().toLocaleString()}\n${'='.repeat(50)}\nDecagon Tickets: ${fmt.num(m.decagonTickets)}\nDecagon Interactions: ${fmt.num(m.decagonInteractions)}\nDecagon FCR: ${fmt.pct(m.fcrRate)}\nContainment Rate: ${fmt.pct(m.containmentRate)}\nCS Assisted Rate: ${fmt.pct(m.csAssistedRate)}\nCompliance Rate: ${fmt.pct(m.complianceRate)}\nStatus Not Closed: ${fmt.num(m.statusNotClosed)}\nMissing Sub Reason: ${fmt.num(m.missingSubReason)}\n`;
  const a=document.createElement('a');a.href=URL.createObjectURL(new Blob([text],{type:'text/plain'}));a.download='way_decagon_summary.txt';a.click();
  showToast('Summary exported','success');
}

// ── SIDEBAR & TAB NAV ──
function setupNav(){
  const TITLES={'upload':'Data Source','validation':'Data Validation','kpis':'Executive KPIs','effectiveness':'AI Effectiveness','compliance':'Decagon Compliance','defects':'System Defects','reasons':'Reason Analysis','tickets':'Master Tickets','executive':'Executive Summary'};
  document.querySelectorAll('.nav-item').forEach(item=>{
    item.addEventListener('click',()=>{
      document.querySelectorAll('.nav-item').forEach(n=>n.classList.remove('active'));
      document.querySelectorAll('.tab-panel').forEach(p=>p.classList.remove('active'));
      item.classList.add('active');
      const tab=item.dataset.tab;
      document.getElementById('tab-'+tab)?.classList.add('active');
      document.getElementById('topbarTitle').textContent=TITLES[tab]||tab;
    });
  });

  document.getElementById('sidebarToggle').addEventListener('click',()=>{
    document.getElementById('sidebar').classList.toggle('collapsed');
    document.querySelector('.main-wrapper').classList.toggle('expanded');
  });
}

// ── INIT ──
document.addEventListener('DOMContentLoaded',()=>{
  setupUpload();
  setupNav();

  document.getElementById('exportPdfBtn').addEventListener('click',exportPDF);
  document.getElementById('exportSummaryBtn').addEventListener('click',exportSummary);
  document.getElementById('applyDateBtn').addEventListener('click',applyDateFilter);
  document.getElementById('clearDateBtn').addEventListener('click',clearDateFilter);
  document.getElementById('closeTimelineModal').addEventListener('click',()=>{document.getElementById('timelineModal').style.display='none';});
  document.getElementById('timelineModal').addEventListener('click',e=>{if(e.target===document.getElementById('timelineModal'))document.getElementById('timelineModal').style.display='none';});
  document.getElementById('recalcDefectsBtn').addEventListener('click',()=>{
    CONFIG.DEFECT_THRESHOLD_SEC=parseInt(document.getElementById('defectThreshold').value)||60;
    computeShortIntervalDefects(STATE.ticketMap);
    STATE.filteredTickets.forEach(t=>{});// re-enrich not needed, shortIntervalFlag set at map level
    renderDefectSection(computeMetrics(STATE.filteredTickets));
    showToast(`Threshold updated to ${CONFIG.DEFECT_THRESHOLD_SEC}s`,'info');
  });

  document.querySelectorAll('.btn-img-export').forEach(btn=>{
    btn.addEventListener('click',e=>{
      e.stopPropagation();
      const chart=STATE.charts[btn.dataset.chart];
      if(!chart)return;
      const a=document.createElement('a');a.download=btn.dataset.chart+'.png';a.href=chart.toBase64Image();a.click();
      showToast('Chart saved as PNG','success');
    });
  });

  document.addEventListener('keydown',e=>{if(e.key==='Escape')document.getElementById('timelineModal').style.display='none';});
  showToast('Dashboard ready — upload your CS All Tickets XLSX or CSV','info',5000);
});

window.showTimeline=showTimeline;
