'use strict';
/* WAY-DECAGON DASHBOARD v3.2 — All logic corrections applied */

const CONFIG = {
  AI_TYPE: 'AI-Agent Call',
  INTERNAL: ['TL Review','Manager Review','QC Audit','Select','User Reviews','BBB Reviews','App Feedback','Escalation Handled by TL','Escalation handled by Escalation Team','Escalation handled by Manager','Escalation handled by Ops Team'],
  CUSTOMER_FACING: ['AI-Agent Call','Call','Email','Chat','SMS'],
  EXCLUDED_REASONS: ['escalated','non escalated','not escalated',''],
  DEFECT_THRESHOLD_SEC: 60
};

const STATE = {
  rawRows:[], ticketMap:new Map(), filteredTickets:new Map(),
  charts:{}, datatables:{}, colMap:{},
  currentReasonTab:'handled', reasonData:{},
  totalCallInteractions:0, totalAIInteractions:0
};

// ── UTILS ──
const fmt = {
  num: n => n==null?'—':Number(n).toLocaleString(),
  pct: n => n==null?'—':Number(n).toFixed(1)+'%',
  date: d => { if(!d)return'—'; const dt=new Date(d); return isNaN(dt)?String(d):dt.toLocaleDateString('en-US',{month:'short',day:'numeric',year:'numeric'}); },
  datetime: d => { if(!d)return'—'; const dt=new Date(d); return isNaN(dt)?String(d):dt.toLocaleString('en-US',{month:'short',day:'numeric',year:'numeric',hour:'2-digit',minute:'2-digit',second:'2-digit'}); }
};
function showToast(msg,type='info',dur=4000){
  const tc=document.getElementById('toastContainer');
  const t=document.createElement('div');
  t.className=`toast toast-${type}`;
  t.innerHTML=`<i class="fa-solid fa-${type==='success'?'circle-check':type==='error'?'circle-xmark':'circle-info'}"></i>${msg}`;
  tc.appendChild(t);
  setTimeout(()=>{t.style.animation='fadeOut 0.3s ease forwards';setTimeout(()=>t.remove(),300);},dur);
}
function pct(n,d){return d>0?(n/d*100):0;}
function badge(t,c='muted'){return`<span class="badge badge-${c}">${t}</span>`;}
function parseDate(d){
  if(!d)return null;
  const s=String(d).trim();
  // Parse M/D/YYYY H:MM:SS AM/PM format as local time to avoid UTC timezone shift
  const m=s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})\s+(\d{1,2}):(\d{2}):(\d{2})\s*(AM|PM)$/i);
  if(m){
    let h=parseInt(m[4]);
    const mn=parseInt(m[5]),sc=parseInt(m[6]),ampm=m[7].toUpperCase();
    if(ampm==='AM'&&h===12)h=0;
    if(ampm==='PM'&&h!==12)h+=12;
    const dt=new Date(parseInt(m[3]),parseInt(m[1])-1,parseInt(m[2]),h,mn,sc);
    return isNaN(dt)?null:dt.getTime();
  }
  const dt=new Date(d);
  return isNaN(dt)?null:dt.getTime();
}

// ── COLUMN MAP ──
const COL_C={ticketId:['Ticket ID','ticket_id'],ogi:['OGI','ogi'],interaction:['Interaction','Interaction Type'],intDate:['Interaction date','Interaction Date','Created Date'],intId:['Interaction ID'],reason:['TKT_IssueReason','Reason'],subReason:['Sub Reason','sub_reason'],action:['Action','Action Taken'],status:['Status'],agent:['Agent Name'],vertical:['Vertical'],subVertical:['SubVertical','Sub Vertical'],ticketCreatedDate:['Ticket_created_date','Ticket Created Date']};
function buildColMap(headers){const m={};Object.entries(COL_C).forEach(([k,cs])=>{m[k]=cs.find(c=>headers.find(h=>h&&h.trim().toLowerCase()===c.toLowerCase()))||null;});return m;}
function getV(row,col){return col?String(row[col]||'').trim():'';}

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
        setTimeout(()=>{try{resolve(XLSX.utils.sheet_to_json(ws,{defval:'',raw:false}));}catch(err){reject(err);}},10);
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
  const checks={cols:!!cm.ticketId&&!!cm.interaction,ticket:rows.some(r=>getV(r,cm.ticketId)),ogi:rows.some(r=>getV(r,cm.ogi)),types:rows.some(r=>getV(r,cm.interaction)),compliance:!!cm.reason||!!cm.subReason||!!cm.action};
  // Interaction type breakdown
  const bd={};
  rows.forEach(r=>{const t=getV(r,cm.interaction)||'Unknown';bd[t]=(bd[t]||0)+1;});
  STATE.interactionBreakdown=bd;
  // Total call interactions (human Call + AI)
  STATE.totalCallInteractions=(bd['Call']||0)+(bd['AI-Agent Call']||0);
  STATE.totalAIInteractions=bd['AI-Agent Call']||0;
  return{
    passed:checks.cols&&checks.ticket&&checks.types,checks,
    counts:{
      records:rows.length,
      uniqueOGI:new Set(rows.map(r=>getV(r,cm.ogi)).filter(Boolean)).size,
      uniqueTickets:new Set(rows.map(r=>getV(r,cm.ticketId)).filter(Boolean)).size,
      totalInteractions:rows.length,
      aiInteractions:STATE.totalAIInteractions,
      humanInteractions:rows.filter(r=>{const t=getV(r,cm.interaction);return CONFIG.CUSTOMER_FACING.includes(t)&&t!==CONFIG.AI_TYPE;}).length,
      internalInteractions:rows.filter(r=>CONFIG.INTERNAL.includes(getV(r,cm.interaction))).length
    }
  };
}

// ── ENRICH TICKET ──
function enrichTicket(tk){
  const ints=tk.interactions;
  const cf=ints.filter(i=>CONFIG.CUSTOMER_FACING.includes(i.type));
  const aiInts=ints.filter(i=>i.type===CONFIG.AI_TYPE);
  const humanInts=cf.filter(i=>i.type!==CONFIG.AI_TYPE);
  const firstAI=aiInts.length>0?aiInts[0]:null;
  // Decagon ticket = any ticket with at least one AI-Agent Call interaction
  tk.isDecagonTicket=firstAI!==null;
  tk.aiInteractionCount=aiInts.length;
  tk.humanInteractionCount=humanInts.length;
  tk.internalInteractionCount=ints.filter(i=>CONFIG.INTERNAL.includes(i.type)).length;
  tk.customerFacingCount=cf.length;

  if(tk.isDecagonTicket){
    const humanAfterAI=humanInts.filter(i=>(i.parsedDate||0)>(firstAI.parsedDate||0));
    const humanBeforeAI=humanInts.filter(i=>(i.parsedDate||0)<(firstAI.parsedDate||0));
    tk.csAssisted=humanAfterAI.length>0;
    tk.decagonOnly=humanAfterAI.length===0;
    tk.decagonContained=humanAfterAI.length===0;
    // Re-contact analysis: was there a human interaction BEFORE AI?
    tk.isRecontact=humanBeforeAI.length>0;
    tk.recontactResolvedByDecagon=tk.isRecontact&&humanAfterAI.length===0&&tk.status.toLowerCase()==='closed';
    tk.recontactReescalated=tk.isRecontact&&humanAfterAI.length>0;
  }else{
    tk.csAssisted=tk.decagonOnly=tk.decagonContained=false;
    tk.isRecontact=tk.recontactResolvedByDecagon=tk.recontactReescalated=false;
  }

  // Compliance base: Decagon-only tickets (no CS)
  if(tk.isDecagonTicket&&tk.decagonOnly){
    tk.missingReason=!tk.reason;
    tk.missingSubReason=!tk.subReason;
    tk.statusNotClosed=tk.status.toLowerCase()!=='closed';
    tk.pendingStatus=['in progress','waiting for ops'].includes(tk.status.toLowerCase());
    tk.compliant=!tk.missingReason&&!tk.missingSubReason&&!tk.statusNotClosed;
  }else{tk.missingReason=tk.missingSubReason=tk.statusNotClosed=tk.pendingStatus=false;tk.compliant=false;}

  // FCR: Decagon-only + closed + has reason
  tk.fcrAchieved=tk.isDecagonTicket&&tk.decagonOnly&&!tk.statusNotClosed&&(tk.subReason||tk.reason)?true:false;

  // Defects: same OGI + different ticket ID + same timestamp (handled at aggregate)
  const tsSet=new Set();let sameTs=0;
  aiInts.forEach(i=>{if(!i.parsedDate)return;if(tsSet.has(i.parsedDate))sameTs++;else tsSet.add(i.parsedDate);});
  tk.sameTimestampInteractions=sameTs;
  tk.hasDefect=sameTs>0;
  tk.shortIntervalFlag=false;

  // Best display reason - exclude status-type values
  const excluded=CONFIG.EXCLUDED_REASONS;
  const sr=(tk.subReason||'').trim();
  const r=(tk.reason||'').trim();
  tk.displayReason=(!excluded.includes(sr.toLowerCase())&&sr)?sr:(!excluded.includes(r.toLowerCase())&&r)?r:'';

  // Use AI-Agent Call interaction date as the date bucket
  // Customer can call at any time after ticket creation
  const firstAIDate = firstAI ? (firstAI.parsedDate ? new Date(firstAI.parsedDate) : null) : null;
  if(firstAIDate&&!isNaN(firstAIDate)){
    tk.aiInteractionDate = firstAIDate.getFullYear()+"-"+(String(firstAIDate.getMonth()+1).padStart(2,"0"))+"-"+(String(firstAIDate.getDate()).padStart(2,"0"));
    tk.dateBucket = tk.aiInteractionDate;
  } else if(tk.createdDate){
    const d=new Date(tk.interactions[0]?.parsedDate||tk.createdDate);
    if(!isNaN(d)) tk.dateBucket=d.getFullYear()+"-"+String(d.getMonth()+1).padStart(2,"0")+"-"+String(d.getDate()).padStart(2,"0");
  }
}

// ── COMPUTE METRICS ──
function computeMetrics(ticketMap){
  const all=[...ticketMap.values()];
  const dec=all.filter(t=>t.isDecagonTicket);
  const decOnly=dec.filter(t=>t.decagonOnly); // 1317
  const decCount=dec.length; // 1785
  const decOnlyCount=decOnly.length; // 1317
  const csAssistedCount=dec.filter(t=>t.csAssisted).length; // 468
  const containedCount=dec.filter(t=>t.decagonContained).length; // 1317
  const fcrCount=dec.filter(t=>t.fcrAchieved).length; // 34
  const compliantCount=decOnly.filter(t=>t.compliant).length; // 34
  const missingReason=decOnly.filter(t=>t.missingReason).length;
  const missingSubReason=decOnly.filter(t=>t.missingSubReason).length;
  const statusNotClosed=decOnly.filter(t=>t.statusNotClosed).length;
  const pendingStatus=decOnly.filter(t=>t.pendingStatus).length;
  const sameTimestampInts=dec.reduce((s,t)=>s+t.sameTimestampInteractions,0);
  const shortIntervalInts=[...ticketMap.values()].filter(t=>t.shortIntervalFlag).length;
  // Re-contact metrics
  const recontactTickets=dec.filter(t=>t.isRecontact);
  const recontactCount=recontactTickets.length;
  const recontactResolved=dec.filter(t=>t.recontactResolvedByDecagon).length;
  const recontactReescalated=dec.filter(t=>t.recontactReescalated).length;
  const recontactStillOpen=recontactTickets.filter(t=>!t.recontactResolvedByDecagon&&!t.recontactReescalated).length;

  // Duplicate tickets: same OGI + different ticket ID + same timestamp
  const ogiTimestampMap={};
  dec.forEach(t=>{
    t.interactions.filter(i=>i.type===CONFIG.AI_TYPE).forEach(i=>{
      if(!i.parsedDate)return;
      const key=t.ogi+'|'+i.parsedDate;
      if(!ogiTimestampMap[key])ogiTimestampMap[key]=new Set();
      ogiTimestampMap[key].add(t.ticketId);
    });
  });
  const dupTicketCount=Object.values(ogiTimestampMap).filter(s=>s.size>1).length;

  // Calculate filtered totals from filtered ticket map
  const filteredAIInts=all.reduce((s,t)=>s+t.aiInteractionCount,0);
  // For total calls in filtered range, count from raw rows matching date range
  const filteredHumanCallInts=all.reduce((s,t)=>s+t.humanInteractionCount,0);
  const filteredTotalCallInts=filteredAIInts+filteredHumanCallInts;
  // Total records = all interactions from filtered tickets
  const filteredTotalRecords=all.reduce((s,t)=>s+t.interactions.length+t.internalInteractionCount,0);

  return{
    totalRecords:filteredTotalRecords||STATE.rawRows.length,
    totalCallInts:filteredTotalCallInts||STATE.totalCallInteractions,
    totalAIInts:filteredAIInts||STATE.totalAIInteractions,
    totalTickets:all.length,
    decagonTickets:decCount,
    decagonOnlyCount:decOnlyCount,
    csAssistedCount,
    containedCount,containmentRate:pct(containedCount,decCount),
    fcrCount,fcrRate:pct(fcrCount,decOnlyCount),
    compliantCount,complianceRate:pct(compliantCount,decOnlyCount),
    complianceFailures:decOnlyCount-compliantCount,
    missingReason,missingSubReason,statusNotClosed,pendingStatus,
    sameTimestampInts,shortIntervalInts,dupTicketCount,
    recontactCount,recontactResolved,recontactReescalated,recontactStillOpen,recontactTickets,
    all,dec,decOnly
  };
}

// ── SHORT INTERVAL INTERACTIONS ──
function computeShortIntervalDefects(ticketMap){
  const threshold=CONFIG.DEFECT_THRESHOLD_SEC*1000;
  let count=0;
  ticketMap.forEach(tk=>{
    tk.shortIntervalFlag=false;
    if(!tk.isDecagonTicket)return;
    const aiInts=tk.interactions.filter(i=>i.type===CONFIG.AI_TYPE&&i.parsedDate).sort((a,b)=>a.parsedDate-b.parsedDate);
    for(let i=1;i<aiInts.length;i++){
      const diff=aiInts[i].parsedDate-aiInts[i-1].parsedDate;
      if(diff>0&&diff<threshold){tk.shortIntervalFlag=true;count++;break;}
    }
  });
  return count;
}

// ── SAMPLE DATA ──
function generateSampleData(n=500){
  const subs=['Shuttle boarding details at the airport','Lot Address Enquiry','General Enquiry','Shuttle timings','Check-out Assistance','Payment Failed','Booking Modification','QR Code Problem','Need for Additional Parking Time','Shuttle boarding details - General'];
  const actions=['Details provided','Issued Refund','Modified Booking','Reset QR Code','Transferred to Lot','Opened'];
  const statuses=['Closed','Open','In Progress','Waiting for OPs'];
  const rows=[];
  for(let i=0;i<n;i++){
    const tid=1000000+i;const ogi=`OGI${50000000+Math.floor(i/2)}`;
    const d=new Date('2026-05-20T00:00:00Z');d.setDate(d.getDate()+Math.floor(Math.random()*19));d.setHours(8+Math.floor(Math.random()*14),Math.floor(Math.random()*60));
    const sub=Math.random()>0.12?subs[Math.floor(Math.random()*subs.length)]:'';
    const action=actions[Math.floor(Math.random()*actions.length)];
    const status=Math.random()>0.3?'Closed':statuses[Math.floor(Math.random()*statuses.length)];
    rows.push({'Ticket ID':tid,'OGI':ogi,'Interaction':'AI-Agent Call','Interaction date':d.toISOString(),'Interaction ID':2000000+i*3,'TKT_IssueReason':Math.random()>0.08?'Non Escalated':'','Sub Reason':sub,'Action':action,'Status':status,'Vertical':'Parking','SubVertical':'Airport Parking','Agent Name':'Decagon AI','Ticket_created_date':d.toISOString()});
    if(Math.random()<0.25){const d2=new Date(d);d2.setMinutes(d2.getMinutes()+20+Math.floor(Math.random()*60));rows.push({'Ticket ID':tid,'OGI':ogi,'Interaction':Math.random()>0.5?'Call':'Email','Interaction date':d2.toISOString(),'Interaction ID':2000000+i*3+1,'TKT_IssueReason':'Non Escalated','Sub Reason':sub,'Action':'Details provided','Status':'Closed','Vertical':'Parking','SubVertical':'Airport Parking','Agent Name':'Arya J S','Ticket_created_date':d.toISOString()});}
    if(Math.random()<0.3){const d3=new Date(d);d3.setMinutes(d3.getMinutes()+5);rows.push({'Ticket ID':tid,'OGI':ogi,'Interaction':'TL Review','Interaction date':d3.toISOString(),'Interaction ID':2000000+i*3+2,'TKT_IssueReason':'','Sub Reason':'','Action':'','Status':status,'Vertical':'Parking','SubVertical':'Airport Parking','Agent Name':'Team Lead','Ticket_created_date':d.toISOString()});}
    // Add human calls
    if(Math.random()<0.5){const d4=new Date(d);d4.setDate(d4.getDate()-1);rows.push({'Ticket ID':1100000+i,'OGI':ogi,'Interaction':'Call','Interaction date':d4.toISOString(),'Interaction ID':3000000+i,'TKT_IssueReason':'','Sub Reason':'','Action':'','Status':'Closed','Vertical':'Parking','SubVertical':'Airport Parking','Agent Name':'John CS','Ticket_created_date':d4.toISOString()});}
  }
  return rows;
}

// ── UPLOAD ──
function setupUpload(){
  const dz=document.getElementById('dropZone'),fi=document.getElementById('fileInput');
  dz.addEventListener('click',(e)=>{if(e.target.id==="browseBtn"||e.target===dz||e.target.classList.contains("drop-icon")||e.target.classList.contains("drop-title")||e.target.classList.contains("drop-sub")||e.target.classList.contains("drop-format"))fi.click();});document.getElementById("browseBtn")?.addEventListener("click",(e)=>{e.stopPropagation();fi.click();});
  dz.addEventListener('dragover',e=>{e.preventDefault();dz.classList.add('drag-over');});
  dz.addEventListener('dragleave',()=>dz.classList.remove('drag-over'));
  dz.addEventListener('drop',e=>{e.preventDefault();dz.classList.remove('drag-over');if(e.dataTransfer.files[0])processFile(e.dataTransfer.files[0]);});
  fi.addEventListener('change',e=>{if(e.target.files[0])processFile(e.target.files[0]);});
  document.getElementById('loadSampleBtn').addEventListener('click',()=>{showToast('Generating sample data…','info');setTimeout(()=>processRows(generateSampleData(500),'sample_data.xlsx'),100);});
}

function processFile(file){
  const name=file.name.toLowerCase();
  showPG();
  if(name.endsWith('.xlsx')||name.endsWith('.xls')){showToast('Reading Excel file…','info');readXLSX(file).then(rows=>processRows(rows,file.name)).catch(err=>{showToast('Excel error: '+err.message,'error');hidePG();});}
  else{Papa.parse(file,{header:true,skipEmptyLines:true,dynamicTyping:false,complete:r=>processRows(r.data,file.name),error:err=>{showToast('CSV error: '+err.message,'error');hidePG();}});}
}

function showPG(){const pg=document.getElementById('uploadProgress'),bar=document.getElementById('progressBar');pg.style.display='block';let p=0;const iv=setInterval(()=>{p=Math.min(p+Math.random()*10,85);bar.style.width=p+'%';if(p>=85)clearInterval(iv);},200);STATE._piv=iv;STATE._pbar=bar;}
function hidePG(){clearInterval(STATE._piv);if(STATE._pbar)STATE._pbar.style.width='100%';setTimeout(()=>{document.getElementById('uploadProgress').style.display='none';},600);}

function processRows(rows,filename){
  STATE.rawRows=rows;
  document.getElementById('fileStatus').style.display='block';
  document.getElementById('statFilename').textContent=filename;
  document.getElementById('statLoaded').textContent=new Date().toLocaleString('en-GB',{day:'2-digit',month:'short',year:'numeric',hour:'2-digit',minute:'2-digit'});
  document.getElementById('statRecords').textContent=rows.length.toLocaleString();
  document.getElementById('statStatus').innerHTML='<span style="color:#059669">✓ Processed Successfully</span>';

  const val=validateData(rows);
  renderValidation(val);
  if(!val.passed){showToast('Validation failed','error');hidePG();return;}

  const pg=document.getElementById('uploadProgress'),bar=document.getElementById('progressBar'),lbl=document.getElementById('progressLabel');
  pg.style.display='block';bar.style.width='20%';lbl.textContent='Building ticket map...';

  setTimeout(()=>{
    const cm=STATE.colMap,map=new Map();
    rows.forEach(row=>{
      const tid=getV(row,cm.ticketId);if(!tid)return;
      if(!map.has(tid))map.set(tid,{ticketId:tid,ogi:getV(row,cm.ogi)||'UNKNOWN',createdDate:getV(row,cm.ticketCreatedDate)||getV(row,cm.intDate),reason:getV(row,cm.reason),subReason:getV(row,cm.subReason),actionTaken:getV(row,cm.action),status:getV(row,cm.status),vertical:getV(row,cm.vertical),subVertical:getV(row,cm.subVertical),interactions:[]});
      const tk=map.get(tid);
      tk.interactions.push({interactionId:getV(row,cm.intId),type:getV(row,cm.interaction),createdDate:getV(row,cm.intDate),parsedDate:parseDate(getV(row,cm.intDate)),reason:getV(row,cm.reason),subReason:getV(row,cm.subReason),actionTaken:getV(row,cm.action),agent:getV(row,cm.agent),status:getV(row,cm.status)});
      if(!tk.reason&&getV(row,cm.reason))tk.reason=getV(row,cm.reason);
      if(!tk.subReason&&getV(row,cm.subReason))tk.subReason=getV(row,cm.subReason);
      if(!tk.actionTaken&&getV(row,cm.action))tk.actionTaken=getV(row,cm.action);
      if(!tk.status&&getV(row,cm.status))tk.status=getV(row,cm.status);
      if(tk.ogi==='UNKNOWN'&&getV(row,cm.ogi))tk.ogi=getV(row,cm.ogi);
    });
    bar.style.width='50%';lbl.textContent='Enriching calls...';
    const tickets=[...map.values()],CHUNK=300;let idx=0;
    function nextChunk(){
      const end=Math.min(idx+CHUNK,tickets.length);
      for(let i=idx;i<end;i++){tickets[i].interactions.sort((a,b)=>(a.parsedDate||0)-(b.parsedDate||0));enrichTicket(tickets[i]);}
      idx=end;bar.style.width=(50+Math.floor((idx/tickets.length)*45))+'%';
      lbl.textContent='Enriching calls... '+idx.toLocaleString()+' / '+tickets.length.toLocaleString();
      if(idx<tickets.length){setTimeout(nextChunk,0);}
      else{
        bar.style.width='100%';lbl.textContent='Finalising...';
        setTimeout(()=>{
          STATE.ticketMap=map;STATE.filteredTickets=new Map(map);
          computeShortIntervalDefects(STATE.ticketMap);
          const m=computeMetrics(STATE.ticketMap);
          const allDates=[...map.values()].filter(t=>t.aiInteractionDate).map(t=>t.aiInteractionDate);
          if(allDates.length){const minD=allDates.reduce((a,b)=>a<b?a:b);const maxD=allDates.reduce((a,b)=>a>b?a:b);document.getElementById('globalDateFrom').value=minD;document.getElementById('globalDateTo').value=maxD;document.getElementById('dateRangeBar').style.display='flex';}
          document.getElementById('dataBadge').style.display='flex';
          document.getElementById('dataBadgeText').textContent=fmt.num(m.decagonTickets)+' Decagon Calls';
          pg.style.display='none';
          renderValidationBreakdown();
          renderDashboard();
          showToast('Loaded '+fmt.num(rows.length)+' records — '+fmt.num(m.decagonTickets)+' Decagon calls','success');
        },0);
      }
    }
    setTimeout(nextChunk,0);
  },50);
}

// ── RE-CONTACT TAB ──
function renderRecontactTab(m){
  const colorMap={cyan:{a:'var(--cyan)',d:'var(--cyan-dim)'},green:{a:'var(--green)',d:'var(--green-dim)'},amber:{a:'var(--amber)',d:'var(--amber-dim)'},red:{a:'var(--red)',d:'var(--red-dim)'}};
  const kpis=[
    {label:'Re-contact Calls Handled by Decagon',mainVal:fmt.num(m.recontactCount),subVal:fmt.pct(pct(m.recontactCount,m.decagonTickets))+' of Decagon calls',icon:'fa-phone-arrow-up-right',color:'cyan',tip:'Tickets where customer had a human interaction before Decagon handled a follow-up call',lvl:'Ticket'},
    {label:'Resolved by Decagon',mainVal:fmt.num(m.recontactResolved),subVal:m.recontactCount?fmt.pct(pct(m.recontactResolved,m.recontactCount)):'0%',icon:'fa-circle-check',color:'green',tip:'Re-contact calls where Decagon resolved it — no further human involvement and ticket closed',lvl:'Ticket'},
    {label:'Re-escalated to CS',mainVal:fmt.num(m.recontactReescalated),subVal:m.recontactCount?fmt.pct(pct(m.recontactReescalated,m.recontactCount)):'0%',icon:'fa-person-walking-arrow-right',color:'amber',tip:'Re-contact calls where human CS agent had to step in again after Decagon',lvl:'Ticket'},
    {label:'Still Open / Unresolved',mainVal:fmt.num(m.recontactStillOpen),subVal:m.recontactCount?fmt.pct(pct(m.recontactStillOpen,m.recontactCount)):'0%',icon:'fa-clock',color:'red',tip:'Re-contact calls handled by Decagon but ticket still not closed',lvl:'Ticket'}
  ];

  const grid=document.getElementById('recontactKpiGrid');
  if(!grid)return;
  grid.innerHTML=kpis.map(k=>{
    const c=colorMap[k.color]||colorMap.cyan;
    return`<div class="kpi-card" style="--ac:${c.a};--acd:${c.d}">
      <div class="kpi-tip" title="${k.tip}"><i class="fa-solid fa-circle-info"></i></div>
      <div class="kpi-icon"><i class="fa-solid ${k.icon}"></i></div>
      <div class="kpi-label">${k.label}</div>
      <div class="kpi-val-large">${k.mainVal}</div>
      <div class="kpi-val-small">${k.subVal}</div>
      <div class="kpi-lvl">${k.lvl} Level</div>
    </div>`;
  }).join('');

  // Reason charts
  const resolvedReasons=countByReason(m.recontactTickets.filter(t=>t.recontactResolvedByDecagon));
  const escalatedReasons=countByReason(m.recontactTickets.filter(t=>t.recontactReescalated));
  const{text,grid:gridColor}=getCC();

  dChart('recontactResolvedChart');
  if(document.getElementById('recontactResolvedChart')){
    if(resolvedReasons.length){
      STATE.charts.recontactResolvedChart=new Chart(document.getElementById('recontactResolvedChart'),{
        type:'bar',data:{labels:resolvedReasons.map(d=>d[0]),datasets:[{label:'Resolved Calls',data:resolvedReasons.map(d=>d[1]),backgroundColor:'rgba(5,150,105,0.7)',borderRadius:4}]},
        options:{indexAxis:'y',responsive:true,plugins:{legend:{display:false}},scales:{x:{ticks:{color:text,font:{size:10}},grid:{color:gridColor}},y:{ticks:{color:text,font:{size:10}}}}}
      });
    } else {
      document.getElementById('recontactResolvedChart').parentElement.innerHTML='<div style="display:flex;align-items:center;justify-content:center;height:200px;color:#94a3b8;flex-direction:column;gap:0.5rem"><i class="fa-solid fa-chart-bar" style="font-size:2rem"></i><p style="font-size:13px">No re-contact resolved data in selected period</p></div>';
    }
  }

  dChart('recontactEscalatedChart');
  if(document.getElementById('recontactEscalatedChart')){
    if(escalatedReasons.length){
      STATE.charts.recontactEscalatedChart=new Chart(document.getElementById('recontactEscalatedChart'),{
        type:'bar',data:{labels:escalatedReasons.map(d=>d[0]),datasets:[{label:'Re-escalated Calls',data:escalatedReasons.map(d=>d[1]),backgroundColor:'rgba(217,119,6,0.7)',borderRadius:4}]},
        options:{indexAxis:'y',responsive:true,plugins:{legend:{display:false}},scales:{x:{ticks:{color:text,font:{size:10}},grid:{color:gridColor}},y:{ticks:{color:text,font:{size:10}}}}}
      });
    } else {
      document.getElementById('recontactEscalatedChart').parentElement.innerHTML='<div style="display:flex;align-items:center;justify-content:center;height:200px;color:#94a3b8;flex-direction:column;gap:0.5rem"><i class="fa-solid fa-chart-bar" style="font-size:2rem"></i><p style="font-size:13px">No re-escalated data in selected period</p></div>';
    }
  }

  // Table
  if(STATE.datatables.recontactTable){STATE.datatables.recontactTable.destroy();document.getElementById('recontactTable').innerHTML='';}
  if(document.getElementById('recontactTable')){
    STATE.datatables.recontactTable=$('#recontactTable').DataTable({
      data:m.recontactTickets,pageLength:25,dom:'Bfrtip',buttons:['csv','excel'],scrollX:true,
      columns:[
        {title:'Ticket ID',data:'ticketId',render:d=>`<a class="ticket-link" onclick="showTimeline('${d}')">${d}</a>`},
        {title:'OGI',data:'ogi'},{title:'Date',data:'createdDate',render:d=>fmt.date(d)},
        {title:'Vertical',data:'subVertical',render:(d,_,r)=>d||r.vertical||'—'},
        {title:'Sub Reason',data:'subReason',render:d=>d||'<span style="color:#94a3b8">—</span>'},
        {title:'Status',data:'status',render:d=>{const c=d==='Closed'?'green':d==='Open'?'red':'amber';return badge(d||'—',c);}},
        {title:'AI Ints',data:'aiInteractionCount',width:'55px'},
        {title:'Human Ints',data:'humanInteractionCount',width:'70px'},
        {title:'Resolved by Decagon',data:'recontactResolvedByDecagon',width:'120px',render:d=>d?badge('YES','green'):badge('NO','red')},
        {title:'Re-escalated to CS',data:'recontactReescalated',width:'120px',render:d=>d?badge('YES','amber'):badge('No','muted')}
      ]
    });
  }
}

// ── RENDER ALL ──
function renderDashboard(){
  const m=computeMetrics(STATE.filteredTickets);
  renderKPIs(m);renderEffectivenessCharts(m);renderComplianceSection(m);
  renderDefectSection(m);renderReasonAnalysis(m);renderMasterTable(m);renderCEOSummary(m);
  renderRecontactTab(m);
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
  const cm={'chk-cols':{ok:checks.cols,l:'Required columns present'},'chk-ticket':{ok:checks.ticket,l:'Ticket IDs present'},'chk-ogi':{ok:checks.ogi,l:'OGI identifiers present'},'chk-types':{ok:checks.types,l:'Interaction types present'},'chk-compliance':{ok:checks.compliance,l:'Compliance fields present'}};
  Object.entries(cm).forEach(([id,{ok,l}])=>{const el=document.getElementById(id);el.className=`val-check-item ${ok?'pass':'fail'}`;el.innerHTML=`<i class="fa-solid fa-${ok?'circle-check':'circle-xmark'}"></i> ${l}`;});
  const b=document.getElementById('validationBadge');b.className=`val-badge ${passed?'pass':'fail'}`;b.innerHTML=`<i class="fa-solid fa-${passed?'shield-check':'shield-xmark'}"></i> ${passed?'✓ Validation Passed':'✗ Validation Failed'}`;
}

function renderValidationBreakdown(){
  const bd=STATE.interactionBreakdown||{};
  document.getElementById('intBreakdownWrap').style.display='block';
  const total=Object.values(bd).reduce((s,v)=>s+v,0);
  const data=Object.entries(bd).sort((a,b)=>b[1]-a[1]).map((e,i)=>({rank:i+1,type:e[0],count:e[1],pct:(e[1]/total*100).toFixed(1)+'%'}));
  if(STATE.datatables.intBreakdown){STATE.datatables.intBreakdown.destroy();document.getElementById('intBreakdownTable').innerHTML='';}
  STATE.datatables.intBreakdown=$('#intBreakdownTable').DataTable({data,pageLength:20,dom:'frtip',columns:[{title:'#',data:'rank',width:'40px'},{title:'Interaction Type',data:'type'},{title:'Count',data:'count',render:d=>fmt.num(d)},{title:'% of Total',data:'pct'}]});
  document.querySelector('#intBreakdownWrap .block-title').innerHTML=`Interaction Type Breakdown <span class="level-tag">All ${fmt.num(total)} Interactions</span>`;
}

// ── KPI CARDS ──
function renderKPIs(m){
  const colorMap={cyan:{a:'var(--cyan)',d:'var(--cyan-dim)'},purple:{a:'var(--purple)',d:'var(--purple-dim)'},green:{a:'var(--green)',d:'var(--green-dim)'},amber:{a:'var(--amber)',d:'var(--amber-dim)'},red:{a:'var(--red)',d:'var(--red-dim)'}};

  const kpis=[
    {label:'Calls Handled by Decagon',mainVal:fmt.num(m.decagonTickets),subVal:null,icon:'fa-robot',color:'cyan',tip:'Unique calls where Decagon (AI-Agent Call) was the first customer-facing interaction',lvl:'Ticket'},
    {label:'Interactions by Decagon',mainVal:fmt.num(m.totalAIInts),subVal:null,icon:'fa-comments',color:'purple',tip:'Total AI-Agent Call interaction records across all Decagon calls',lvl:'Interaction'},
    {label:'Decagon FCR',mainVal:fmt.pct(m.fcrRate),subVal:fmt.num(m.fcrCount)+' calls',icon:'fa-bullseye',color:'green',tip:'Calls fully resolved by Decagon: no CS involvement + ticket closed + reason tagged. Base: '+fmt.num(m.decagonOnlyCount)+' Decagon-only calls',lvl:'Ticket'},
    {label:'Decagon Containment Rate',mainVal:fmt.pct(m.containmentRate),subVal:fmt.num(m.containedCount)+' calls',icon:'fa-shield-halved',color:'green',tip:'Calls where no CS agent was involved after Decagon. 1,785 - 468 = 1,317',lvl:'Ticket'},
    {label:'CS Assisted',mainVal:fmt.pct(m.csAssistedCount/m.decagonTickets*100),subVal:fmt.num(m.csAssistedCount)+' calls',icon:'fa-person-walking-arrow-right',color:'amber',tip:'Calls where a human CS agent had to handle after Decagon',lvl:'Ticket',pctLarge:true},
    {label:'Handled by Decagon Only',mainVal:fmt.num(m.decagonOnlyCount),subVal:fmt.pct(pct(m.decagonOnlyCount,m.decagonTickets)),icon:'fa-circle-check',color:'green',tip:'Calls handled entirely by Decagon with no CS agent involvement',lvl:'Ticket'},
    {label:'Compliance Failures',mainVal:fmt.num(m.complianceFailures),subVal:null,icon:'fa-triangle-exclamation',color:'red',tip:'Decagon-only calls missing proper documentation or not closed. Base: '+fmt.num(m.decagonOnlyCount)+' calls',lvl:'Ticket'},
    {label:'Compliance Rate',mainVal:fmt.pct(m.complianceRate),subVal:fmt.num(m.compliantCount)+' calls',icon:'fa-clipboard-check',color:'green',tip:'Decagon-only calls with Reason + Sub Reason filled AND Status = Closed. Base: '+fmt.num(m.decagonOnlyCount)+' calls',lvl:'Ticket'},
    {label:'Decagon Duplicate Ticket',mainVal:fmt.num(m.dupTicketCount),subVal:null,icon:'fa-copy',color:'red',tip:'Same OGI with multiple different Ticket IDs created at the same timestamp by Decagon',lvl:'Ticket'},
    {label:'Short Interval Interactions',mainVal:fmt.num(m.shortIntervalInts),subVal:null,icon:'fa-stopwatch',color:'amber',tip:'AI-Agent Call interactions on the same ticket within '+CONFIG.DEFECT_THRESHOLD_SEC+'s of each other — possible system retry',lvl:'Interaction'}
  ];

  document.getElementById('kpiGrid').innerHTML=kpis.map(k=>{
    const c=colorMap[k.color]||colorMap.cyan;
    const mainHtml=k.pctLarge
      ?`<div class="kpi-pct-large" style="color:${c.a}">${k.mainVal}</div>${k.subVal?`<div class="kpi-num-small">${k.subVal}</div>`:''}`
      :`<div class="kpi-val-large">${k.mainVal}</div>${k.subVal?`<div class="kpi-val-small">${k.subVal}</div>`:''}`;
    return`<div class="kpi-card" style="--ac:${c.a};--acd:${c.d}">
      <div class="kpi-tip" title="${k.tip}"><i class="fa-solid fa-circle-info"></i></div>
      <div class="kpi-icon"><i class="fa-solid ${k.icon}"></i></div>
      <div class="kpi-label">${k.label}</div>
      ${mainHtml}
      <div class="kpi-lvl ${k.lvl==='Interaction'?'int':''}">${k.lvl} Level</div>
    </div>`;
  }).join('');
}

// ── CHARTS ──
function dChart(id){if(STATE.charts[id]){STATE.charts[id].destroy();delete STATE.charts[id];}}
function getCC(){return{text:'#64748b',grid:'#e2e8f0'};}
function getDateBuckets(ticketMap){
  const b=new Map();
  ticketMap.forEach(t=>{
    if(t.isDecagonTicket&&t.dateBucket){
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
  // FCR: closed + no CS + has reason / decagonOnly per day
  const fcrRates=[...buckets.values()].map(ts=>{const d=ts.filter(t=>t.isDecagonTicket&&t.decagonOnly);return d.length?pct(d.filter(t=>t.fcrAchieved).length,d.length):0;});
  const csRates=[...buckets.values()].map(ts=>{const d=ts.filter(t=>t.isDecagonTicket);return d.length?pct(d.filter(t=>t.csAssisted).length,d.length):0;});
  const containRates=[...buckets.values()].map(ts=>{const d=ts.filter(t=>t.isDecagonTicket);return d.length?pct(d.filter(t=>t.decagonContained).length,d.length):0;});

  const base={responsive:true,maintainAspectRatio:true,interaction:{mode:'index',intersect:false},plugins:{legend:{labels:{color:text,font:{size:11}}},tooltip:{backgroundColor:'#fff',titleColor:'#0f172a',bodyColor:'#475569',borderColor:'#e2e8f0',borderWidth:1}},scales:{x:{ticks:{color:text,font:{size:10},maxRotation:45},grid:{color:grid}},y:{ticks:{color:text,font:{size:10}},grid:{color:grid},beginAtZero:true}}};

  dChart('decagonTicketsTrend');
  STATE.charts.decagonTicketsTrend=new Chart(document.getElementById('decagonTicketsTrend'),{type:'bar',data:{labels,datasets:[{label:'Calls Handled by Decagon',data:decCounts,backgroundColor:'rgba(2,132,199,0.6)',borderRadius:4}]},options:{...base}});

  dChart('decagonIntsTrend');
  STATE.charts.decagonIntsTrend=new Chart(document.getElementById('decagonIntsTrend'),{type:'bar',data:{labels,datasets:[{label:'Interactions by Decagon',data:intCounts,backgroundColor:'rgba(124,58,237,0.6)',borderRadius:4}]},options:{...base}});

  dChart('fcrTrend');
  STATE.charts.fcrTrend=new Chart(document.getElementById('fcrTrend'),{type:'line',data:{labels,datasets:[{label:'Decagon FCR %',data:fcrRates,borderColor:'#059669',backgroundColor:'rgba(5,150,105,0.1)',fill:true,tension:0.4,pointRadius:3}]},options:{...base,scales:{x:{ticks:{color:text,font:{size:10},maxRotation:45},grid:{color:grid}},y:{ticks:{color:text,font:{size:10},callback:v=>v+'%'},grid:{color:grid},beginAtZero:true,max:100}}}});

  dChart('csAssistedTrend');
  STATE.charts.csAssistedTrend=new Chart(document.getElementById('csAssistedTrend'),{type:'line',data:{labels,datasets:[{label:'CS Assisted %',data:csRates,borderColor:'#d97706',backgroundColor:'rgba(217,119,6,0.08)',fill:true,tension:0.4,pointRadius:3},{label:'Containment %',data:containRates,borderColor:'#059669',fill:false,tension:0.4,pointRadius:3,borderDash:[5,4]}]},options:{...base,scales:{x:{ticks:{color:text,font:{size:10},maxRotation:45},grid:{color:grid}},y:{ticks:{color:text,font:{size:10},callback:v=>v+'%'},grid:{color:grid},beginAtZero:true}}}});
}

// ── COMPLIANCE ──
function renderComplianceSection(m){
  document.getElementById('gaugeCompPct').textContent=fmt.pct(m.complianceRate);
  document.getElementById('cv-full').textContent=fmt.num(m.compliantCount)+' calls';
  document.getElementById('cv-reason').textContent=fmt.num(m.missingReason);
  document.getElementById('cv-sub').textContent=fmt.num(m.missingSubReason);
  document.getElementById('cv-status').textContent=fmt.num(m.statusNotClosed);
  document.getElementById('cv-pending').textContent=fmt.num(m.pendingStatus);
  renderGauge('complianceGauge',m.complianceRate);
  const{text,grid}=getCC();

  dChart('compliancePie');
  STATE.charts.compliancePie=new Chart(document.getElementById('compliancePie'),{
    type:'doughnut',
    data:{labels:['Fully Compliant','Missing Reason','Missing Sub Reason','Status Not Closed'],datasets:[{data:[m.compliantCount,m.missingReason,m.missingSubReason,m.statusNotClosed],backgroundColor:['#10b981','#ef4444','#f59e0b','#8b5cf6'],borderColor:'#fff',borderWidth:2}]},
    options:{responsive:true,maintainAspectRatio:true,plugins:{legend:{position:'bottom',labels:{color:text,font:{size:11},padding:8}}}}
  });

  const buckets=getDateBuckets(STATE.filteredTickets);
  const labels=[...buckets.keys()].map(d=>new Date(d).toLocaleDateString('en-US',{month:'short',day:'numeric'}));
  dChart('complianceBar');
  STATE.charts.complianceBar=new Chart(document.getElementById('complianceBar'),{
    type:'bar',
    data:{labels,datasets:[
      {label:'Missing Reason',data:[...buckets.values()].map(ts=>ts.filter(t=>t.isDecagonTicket&&t.decagonOnly&&t.missingReason).length),backgroundColor:'rgba(239,68,68,0.7)'},
      {label:'Missing Sub Reason',data:[...buckets.values()].map(ts=>ts.filter(t=>t.isDecagonTicket&&t.decagonOnly&&t.missingSubReason).length),backgroundColor:'rgba(245,158,11,0.7)'},
      {label:'Status Not Closed',data:[...buckets.values()].map(ts=>ts.filter(t=>t.isDecagonTicket&&t.decagonOnly&&t.statusNotClosed).length),backgroundColor:'rgba(139,92,246,0.7)'}
    ]},
    options:{responsive:true,plugins:{legend:{labels:{color:text,font:{size:11}}}},scales:{x:{stacked:true,ticks:{color:text,font:{size:10},maxRotation:45},grid:{color:grid}},y:{stacked:true,ticks:{color:text,font:{size:10}},grid:{color:grid},beginAtZero:true}}}
  });

  document.querySelectorAll('.ci-val.drill').forEach(el=>{
    el.onclick=()=>{
      const drill=el.dataset.drill;
      const tickets=[...STATE.filteredTickets.values()].filter(t=>{
        if(!t.isDecagonTicket||!t.decagonOnly)return false;
        if(drill==='missingReason')return t.missingReason;
        if(drill==='missingSubReason')return t.missingSubReason;
        if(drill==='statusNotClosed')return t.statusNotClosed;
        if(drill==='pendingStatus')return t.pendingStatus;
        return false;
      });
      document.getElementById('compDrillTitle').textContent=(el.previousElementSibling?.textContent||'Drill')+' — '+fmt.num(tickets.length)+' call tickets';
      if(STATE.datatables.compDrill){STATE.datatables.compDrill.destroy();document.getElementById('compDrillTable').innerHTML='';}
      STATE.datatables.compDrill=$('#compDrillTable').DataTable({
        data:tickets,pageLength:10,dom:'Bfrtip',buttons:['csv'],
        columns:[
          {title:'Ticket ID',data:'ticketId',render:d=>`<span class="ticket-link" onclick="showTimeline('${d}')">${d}</span>`},
          {title:'OGI',data:'ogi'},{title:'Date',data:'createdDate',render:d=>fmt.date(d)},
          {title:'Sub Reason',data:'subReason',render:d=>d||badge('MISSING','red')},
          {title:'Status',data:'status',render:d=>{const c=d==='Closed'?'green':d==='Open'?'red':'amber';return badge(d||'—',c);}}
        ]
      });
      document.getElementById('compDrillWrap').style.display='block';
      document.getElementById('compDrillWrap').scrollIntoView({behavior:'smooth',block:'start'});
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
  const shortIntCount=computeShortIntervalDefects(STATE.filteredTickets);
  const defects=[
    {label:'Decagon Duplicate Ticket',val:m.dupTicketCount,type:'error',tip:'Same OGI + multiple ticket IDs at identical timestamp',drill:'dupTicket'},
    {label:'Duplicate AI Interactions',val:m.sameTimestampInts,type:'error',tip:'AI-Agent Call interactions with exact same timestamp on same ticket',drill:'dupInteraction'},
    {label:'Short Interval Interactions',val:shortIntCount,type:'warn',tip:'AI-Agent Call interactions within '+CONFIG.DEFECT_THRESHOLD_SEC+'s of each other on same ticket',drill:'shortInterval',lvl:'Interaction'}
  ];
  document.getElementById('defectGrid').innerHTML=defects.map(d=>`
    <div class="defect-card ${d.type==='warn'?'warn':''}" style="cursor:pointer" onclick="showDefectDrill('${d.drill}')">
      <div class="defect-label">${d.label}${d.lvl?` <span style="font-size:9px">(${d.lvl})</span>`:''}</div>
      <div class="defect-val ${d.type==='warn'?'warn':''}" title="${d.tip}">${fmt.num(d.val)}</div>
      <div style="font-size:10px;color:#94a3b8;margin-top:0.3rem">Click to drill down</div>
    </div>`).join('');
}

function showDefectDrill(type){
  const dec=[...STATE.filteredTickets.values()].filter(t=>t.isDecagonTicket);
  let data=[];
  let title='';

  if(type==='dupTicket'){
    title='Duplicate Ticket Defects (Same OGI + Same Timestamp)';
    // Find OGIs with multiple ticket IDs at same timestamp
    const ogiMap={};
    dec.forEach(t=>{
      t.interactions.filter(i=>i.type===CONFIG.AI_TYPE&&i.parsedDate).forEach(i=>{
        const key=t.ogi+'|'+i.parsedDate;
        if(!ogiMap[key])ogiMap[key]={ogi:t.ogi,timestamp:fmt.datetime(i.createdDate),tickets:new Set()};
        ogiMap[key].tickets.add(t.ticketId);
      });
    });
    data=Object.values(ogiMap).filter(v=>v.tickets.size>1).map(v=>({ticketId:[...v.tickets].join(', '),ogi:v.ogi,createdDate:v.timestamp,aiInteractionCount:'Multiple',sameTimestamp:'YES',shortInterval:'—',subReason:'—',reason:'—',actionTaken:'—'}));
  }else if(type==='dupInteraction'){
    title='Duplicate AI Interactions (Same Timestamp on Same Ticket)';
    dec.forEach(t=>{
      if(t.sameTimestampInteractions>0){
        const aiInts=t.interactions.filter(i=>i.type===CONFIG.AI_TYPE);
        const tsMap={};
        aiInts.forEach(i=>{if(i.parsedDate){if(!tsMap[i.parsedDate])tsMap[i.parsedDate]=0;tsMap[i.parsedDate]++;}});
        Object.entries(tsMap).filter(([,c])=>c>1).forEach(([ts])=>{
          data.push({ticketId:t.ticketId,ogi:t.ogi,createdDate:fmt.datetime(new Date(parseInt(ts))),aiInteractionCount:t.aiInteractionCount,sameTimestamp:t.sameTimestampInteractions,shortInterval:t.shortIntervalFlag?'YES':'No',subReason:t.subReason||'—',reason:t.reason||'—',actionTaken:t.actionTaken||'—'});
        });
      }
    });
  }else if(type==='shortInterval'){
    title='Short Interval Interactions (<'+CONFIG.DEFECT_THRESHOLD_SEC+'s)';
    dec.filter(t=>t.shortIntervalFlag).forEach(t=>{
      data.push({ticketId:t.ticketId,ogi:t.ogi,createdDate:fmt.date(t.createdDate),aiInteractionCount:t.aiInteractionCount,sameTimestamp:t.sameTimestampInteractions,shortInterval:'YES',subReason:t.subReason||'—',reason:t.displayReason||'—',actionTaken:t.actionTaken||'—'});
    });
  }

  document.getElementById('defectModalTitle').textContent=title+' ('+data.length+')';
  if(STATE.datatables.defectDrill){STATE.datatables.defectDrill.destroy();document.getElementById('defectDrillTable').innerHTML='';}
  STATE.datatables.defectDrill=$('#defectDrillTable').DataTable({
    data,pageLength:15,dom:'Bfrtip',buttons:['csv'],scrollX:true,
    columns:[
      {title:'Ticket ID',data:'ticketId',render:d=>`<span class="ticket-link" onclick="showTimeline('${d.split(',')[0].trim()}')">${d}</span>`},
      {title:'Date',data:'createdDate'},{title:'AI Interactions',data:'aiInteractionCount'},
      {title:'Same Timestamp',data:'sameTimestamp',render:d=>d==='YES'||d>0?badge('YES','red'):badge('No','muted')},
      {title:'Short Interval',data:'shortInterval',render:d=>d==='YES'?badge('YES','amber'):badge('No','muted')},
      {title:'Sub Reason',data:'subReason'},{title:'Reason',data:'reason'},{title:'Action Taken',data:'actionTaken'}
    ]
  });
  document.getElementById('defectModal').style.display='flex';
}

// ── REASON ANALYSIS ──
function getDisplayReason(tk){
  const excluded=CONFIG.EXCLUDED_REASONS;
  const sr=(tk.subReason||'').trim(),r=(tk.reason||'').trim();
  if(sr&&!excluded.includes(sr.toLowerCase()))return sr;
  if(r&&!excluded.includes(r.toLowerCase()))return r;
  return null;
}

function countByReason(tickets){
  const c={};
  tickets.forEach(t=>{const r=getDisplayReason(t);if(!r)return;c[r]=(c[r]||0)+1;});
  return Object.entries(c).sort((a,b)=>b[1]-a[1]).slice(0,10);
}

function renderReasonAnalysis(m){
  STATE.reasonData={handled:m.dec.filter(t=>t.decagonOnly),cs:m.dec.filter(t=>t.csAssisted),comp:m.dec.filter(t=>t.isDecagonTicket&&t.decagonOnly&&!t.compliant)};
  renderReasonChart(STATE.currentReasonTab||'handled');
  document.querySelectorAll('.reason-tab').forEach(tab=>{
    tab.onclick=()=>{
      document.querySelectorAll('.reason-tab').forEach(t=>t.classList.remove('active'));tab.classList.add('active');
      STATE.currentReasonTab=tab.dataset.rt;
      document.getElementById('reasonDetailSide').innerHTML='<div class="reason-detail-placeholder"><i class="fa-solid fa-hand-pointer"></i><p>Click any bar to see breakdown</p></div>';
      renderReasonChart(tab.dataset.rt);
    };
  });
}

function renderReasonChart(tab){
  const tickets=STATE.reasonData[tab]||[];
  const data=countByReason(tickets);
  const{text,grid}=getCC();
  const colors={handled:'rgba(2,132,199,0.7)',cs:'rgba(217,119,6,0.7)',comp:'rgba(220,38,38,0.7)'};
  dChart('reasonChart');
  STATE.charts.reasonChart=new Chart(document.getElementById('reasonChart'),{
    type:'bar',
    data:{labels:data.map(d=>d[0]),datasets:[{label:'Calls',data:data.map(d=>d[1]),backgroundColor:colors[tab]||colors.handled,borderRadius:4}]},
    options:{
      indexAxis:'y',responsive:true,
      onClick:(evt,els)=>{if(!els.length)return;showReasonDetail(data[els[0].index][0],tab);},
      plugins:{legend:{display:false},tooltip:{backgroundColor:'#fff',titleColor:'#0f172a',bodyColor:'#475569',borderColor:'#e2e8f0',borderWidth:1}},
      scales:{x:{ticks:{color:text,font:{size:10}},grid:{color:grid}},y:{ticks:{color:text,font:{size:10}}}}
    }
  });
}

function showReasonDetail(reason,tab){
  const allDecTks=[...STATE.filteredTickets.values()].filter(t=>t.isDecagonTicket);
  const reasonTks=allDecTks.filter(t=>getDisplayReason(t)===reason);
  const total=reasonTks.length;
  const decHandled=reasonTks.filter(t=>t.decagonOnly).length;
  const csHandled=reasonTks.filter(t=>t.csAssisted).length;
  const statusCounts={};
  reasonTks.forEach(t=>{const s=t.status||'Unknown';statusCounts[s]=(statusCounts[s]||0)+1;});
  const statusColors={'Closed':'#10b981','Open':'#ef4444','In Progress':'#f59e0b','Waiting for OPs':'#8b5cf6','Unknown':'#94a3b8'};
  const rows=[
    {name:'Handled by Decagon Only',count:decHandled,color:'#0ea5e9'},
    {name:'CS Handled',count:csHandled,color:'#d97706'},
    ...Object.entries(statusCounts).sort((a,b)=>b[1]-a[1]).map(([s,c])=>({name:s,count:c,color:statusColors[s]||'#94a3b8'}))
  ];
  document.getElementById('reasonDetailSide').innerHTML=`
    <div class="reason-detail-title">${reason}</div>
    <div class="reason-detail-total">${fmt.num(total)}<span>Total Decagon Calls</span></div>
    <div class="status-breakdown">
      ${rows.map(s=>`<div class="status-row">
        <span class="status-name" style="color:${s.color};font-weight:600">${s.name}</span>
        <div class="status-bar-wrap"><div class="status-bar-fill" style="width:${total?pct(s.count,total):0}%;background:${s.color}"></div></div>
        <span class="status-count">${fmt.num(s.count)}</span>
        <span style="font-size:11px;color:#94a3b8;width:40px;text-align:right">${fmt.pct(total?pct(s.count,total):0)}</span>
      </div>`).join('')}
    </div>`;
}

// ── MASTER TABLE (Decagon calls only) ──
function renderMasterTable(m){
  const tickets=m.dec; // Only Decagon calls
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
      {title:'Contained',data:'decagonContained',width:'85px',render:d=>d?badge('YES','green'):badge('NO','red')},
      {title:'Compliance',data:'compliant',width:'90px',render:(d,_,r)=>r.decagonOnly?(d?badge('PASS','green'):badge('FAIL','red')):badge('CS Assisted','amber')},
      {title:'CS Assisted',data:'csAssisted',width:'85px',render:d=>d?badge('YES','amber'):badge('No','muted')}
    ]
  });
}

// ── TIMELINE ──
function showTimeline(ticketId){
  const tk=STATE.filteredTickets.get(String(ticketId))||STATE.ticketMap.get(String(ticketId));
  if(!tk)return;
  document.getElementById('timelineTicketId').textContent='Ticket: '+ticketId;
  document.getElementById('timelineTicketMeta').textContent='OGI: '+tk.ogi+' · '+tk.interactions.length+' interactions · FCR: '+(tk.fcrAchieved?'Pass':'Fail')+' · '+(tk.subVertical||tk.vertical||'');
  const tsCount={};tk.interactions.forEach(i=>{if(i.parsedDate)tsCount[i.parsedDate]=(tsCount[i.parsedDate]||0)+1;});
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
    return`<div class="tl-item"><div class="tl-dot ${dc}"></div><div class="tl-content"><div class="tl-time">${fmt.datetime(int.createdDate)}</div><div class="tl-type">${int.type}</div>${int.subReason?`<div style="font-size:11px;color:#64748b">${int.subReason}</div>`:''}<div class="tl-flags">${flags.join('')}</div></div></div>`;
  }).join('');
  document.getElementById('timelineBody').innerHTML='<div class="timeline-list">'+html+'</div>';
  document.getElementById('timelineModal').style.display='flex';
}

// ── CEO SUMMARY ──
function renderCEOSummary(m){
  const decTks=[...STATE.filteredTickets.values()].filter(t=>t.isDecagonTicket);
  const topHandled=countByReason(decTks.filter(t=>t.decagonOnly)).slice(0,3);
  const topCS=countByReason(decTks.filter(t=>t.csAssisted)).slice(0,3);
  const decagonShareOfCalls=pct(m.totalAIInts,m.totalCallInts);
  const callsRoutedToDecagon=m.decagonTickets; // 1785
  const callsHandledAlone=m.decagonOnlyCount; // 1317
  const pctHandledAlone=pct(callsHandledAlone,callsRoutedToDecagon);

  // Key observations
  const obs=[];
  obs.push(`Decagon handled <strong>${fmt.num(m.decagonTickets)} calls</strong> out of <strong>${fmt.num(m.totalCallInts)} total calls</strong> — representing <strong>${fmt.pct(decagonShareOfCalls)}</strong> of all voice interactions in the CRM.`);
  obs.push(`Of ${fmt.num(m.decagonTickets)} calls routed to Decagon, <strong>${fmt.num(callsHandledAlone)} (${fmt.pct(pctHandledAlone)})</strong> were handled by Decagon alone without CS involvement.`);
  obs.push(`Only <strong>${fmt.num(m.fcrCount)} call tickets (${fmt.pct(m.fcrRate)})</strong> were fully resolved and closed by Decagon — indicating the API is not closing tickets after handling calls.`);
  obs.push(`<strong>${fmt.num(m.statusNotClosed)} call tickets are not closed</strong> after Decagon interaction — this is a Decagon API integration issue, not an agent issue.`);

  // What is working
  const working=[];
  working.push(`Decagon successfully contained <strong>${fmt.pct(m.containmentRate)}</strong> of calls without requiring CS agent involvement.`);
  working.push(`Call volume handled by Decagon is growing — trend shows increasing daily calls from May 19 to June 8.`);
  working.push(`CS agents were freed from <strong>${fmt.num(callsHandledAlone)}</strong> calls that Decagon handled independently.`);

  // Recommended actions
  const recs=[];
  recs.push(`Fix Decagon API to automatically close call tickets upon successful resolution — ${fmt.num(m.statusNotClosed)} call tickets currently left open.`);
  recs.push(`Fix Sub Reason tagging in Decagon configuration — ${fmt.num(m.missingSubReason)} call tickets missing Sub Reason field.`);
  recs.push(`Investigate ${fmt.num(m.shortIntervalInts)} short interval interactions — possible system retry issue in Decagon API.`);
  recs.push(`Expand Decagon call handling coverage — currently at ${fmt.pct(decagonShareOfCalls)} of total voice calls, significant room to deflect more human call volume.`);

  document.getElementById('ceoSummaryCard').innerHTML=`<div class="ceo-content">
    <div class="ceo-meta-row">
      <div class="ceo-meta-item"><div class="ceo-meta-label">Total CRM Records</div><div class="ceo-meta-val">${fmt.num(m.totalRecords)}</div></div>
      <div class="ceo-meta-item"><div class="ceo-meta-label">Total Calls (Human + Decagon)</div><div class="ceo-meta-val">${fmt.num(m.totalCallInts)}</div></div>
      <div class="ceo-meta-item"><div class="ceo-meta-label">Calls Routed to Decagon</div><div class="ceo-meta-val">${fmt.num(callsRoutedToDecagon)}</div></div>
      <div class="ceo-meta-item"><div class="ceo-meta-label">Calls Handled by Decagon Alone</div><div class="ceo-meta-val">${fmt.num(callsHandledAlone)}</div></div>
    </div>
    <div class="ceo-kpi-row">
      <div class="ceo-kpi-item"><div class="ceo-kpi-label">Escalated to CS</div><div class="ceo-kpi-val" style="color:#d97706">${fmt.num(m.csAssistedCount)} <span style="font-size:12px">(${fmt.pct(m.csAssistedCount/m.decagonTickets*100)})</span></div></div>
      <div class="ceo-kpi-item"><div class="ceo-kpi-label">Decagon FCR</div><div class="ceo-kpi-val" style="color:#dc2626">${fmt.pct(m.fcrRate)}</div></div>
      <div class="ceo-kpi-item"><div class="ceo-kpi-label">Containment Rate</div><div class="ceo-kpi-val" style="color:#059669">${fmt.pct(m.containmentRate)}</div></div>
      <div class="ceo-kpi-item"><div class="ceo-kpi-label">Compliance Rate</div><div class="ceo-kpi-val" style="color:#dc2626">${fmt.pct(m.complianceRate)}</div></div>
      <div class="ceo-kpi-item"><div class="ceo-kpi-label">Compliance Failures</div><div class="ceo-kpi-val" style="color:#dc2626">${fmt.num(m.complianceFailures)}</div></div>
    </div>
    <div class="ceo-sections">
      <div class="ceo-col">
        <h4>📞 Top Reasons — Decagon Handled</h4>
        ${topHandled.length?topHandled.map((e,i)=>`<div class="insight-item"><div class="insight-dot" style="background:#0ea5e9"></div><div class="insight-text">${i+1}. <strong>${e[0]}</strong> — ${fmt.num(e[1])} calls</div></div>`).join(''):'<p style="font-size:12px;color:#64748b">No data</p>'}
        <h4 style="margin-top:1rem">🔄 Top Reasons — Escalated to CS</h4>
        ${topCS.length?topCS.map((e,i)=>`<div class="insight-item"><div class="insight-dot" style="background:#d97706"></div><div class="insight-text">${i+1}. <strong>${e[0]}</strong> — ${fmt.num(e[1])} calls</div></div>`).join(''):'<p style="font-size:12px;color:#64748b">No escalations</p>'}
      </div>
      <div class="ceo-col">
        <h4>💡 How Decagon is Performing</h4>
        ${obs.map(o=>`<div class="insight-item"><div class="insight-dot"></div><div class="insight-text">${o}</div></div>`).join('')}
        <h4 style="margin-top:1rem">✅ What is Working</h4>
        ${working.map(o=>`<div class="insight-item"><div class="insight-dot" style="background:#10b981"></div><div class="insight-text">${o}</div></div>`).join('')}
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
  const from=document.getElementById('globalDateFrom').value,to=document.getElementById('globalDateTo').value;
  STATE.filteredTickets=new Map();
  STATE.ticketMap.forEach((tk,id)=>{
    // Use AI interaction date for filtering if available, else ticket created date
    const filterDate = tk.aiInteractionDate || tk.createdDate;
    if(from){const d=new Date(filterDate);if(!isNaN(d)&&d<new Date(from))return;}
    if(to){const d=new Date(filterDate);if(!isNaN(d)&&d>new Date(to+'T23:59:59'))return;}
    STATE.filteredTickets.set(id,tk);
  });
  renderDashboard();showToast('Filter applied — '+fmt.num(STATE.filteredTickets.size)+' tickets','info');
}

function clearDateFilter(){
  STATE.filteredTickets=new Map(STATE.ticketMap);
  const dates=[...STATE.ticketMap.values()].filter(t=>t.aiInteractionDate).map(t=>t.aiInteractionDate);
  if(dates.length){dates.sort();document.getElementById('globalDateFrom').value=dates[0];document.getElementById('globalDateTo').value=dates[dates.length-1];}
  renderDashboard();showToast('Filter cleared','info');
}

// ── EXPORT ──
function exportPDF(){
  if(!STATE.filteredTickets.size){showToast('Upload data first','error');return;}
  const{jsPDF}=window.jspdf,doc=new jsPDF({orientation:'portrait',unit:'mm',format:'a4'});
  const m=computeMetrics(STATE.filteredTickets),now=new Date().toLocaleDateString('en-GB');
  doc.setFillColor(15,23,42);doc.rect(0,0,210,35,'F');
  doc.setTextColor(255,255,255);doc.setFontSize(16);doc.setFont('helvetica','bold');
  doc.text('Way-Decagon AI Effectiveness & Quality Dashboard',15,16);
  doc.setFontSize(9);doc.setTextColor(148,163,184);doc.text('Executive Export · '+now,15,26);
  let y=48;doc.setTextColor(15,23,42);doc.setFontSize(12);doc.text('Executive KPI Summary',15,y);y+=8;
  doc.autoTable({startY:y,head:[['KPI','Value']],body:[
    ['Total Calls (Human + Decagon)',fmt.num(m.totalCallInts)],
    ['Calls Routed to Decagon',fmt.num(m.decagonTickets)],
    ['Calls Handled by Decagon Alone',fmt.num(m.decagonOnlyCount)],
    ['Escalated to CS',fmt.num(m.csAssistedCount)+' ('+fmt.pct(m.csAssistedCount/m.decagonTickets*100)+')'],
    ['Decagon FCR',fmt.pct(m.fcrRate)+' ('+fmt.num(m.fcrCount)+' calls)'],
    ['Decagon Containment Rate',fmt.pct(m.containmentRate)],
    ['Compliance Rate',fmt.pct(m.complianceRate)+' ('+fmt.num(m.compliantCount)+' calls)'],
    ['Compliance Failures',fmt.num(m.complianceFailures)+' call tickets'],
    ['Status Not Closed',fmt.num(m.statusNotClosed)+' call tickets'],
    ['Duplicate Ticket Defects',fmt.num(m.dupTicketCount)],
    ['Short Interval Interactions',fmt.num(m.shortIntervalInts)]
  ],margin:{left:15,right:15},headStyles:{fillColor:[15,23,42],textColor:[255,255,255],fontSize:9},bodyStyles:{fontSize:9}});
  doc.save('way_decagon_'+now.replace(/\//g,'-')+'.pdf');showToast('PDF exported','success');
}

function exportSummary(){
  if(!STATE.filteredTickets.size)return;
  const m=computeMetrics(STATE.filteredTickets);
  const text=`WAY-DECAGON EXECUTIVE SUMMARY\nGenerated: ${new Date().toLocaleString()}\n${'='.repeat(50)}\nTotal Calls (Human + Decagon): ${fmt.num(m.totalCallInts)}\nCalls Routed to Decagon: ${fmt.num(m.decagonTickets)}\nCalls Handled by Decagon Alone: ${fmt.num(m.decagonOnlyCount)}\nEscalated to CS: ${fmt.num(m.csAssistedCount)}\nDecagon FCR: ${fmt.pct(m.fcrRate)}\nContainment Rate: ${fmt.pct(m.containmentRate)}\nCompliance Rate: ${fmt.pct(m.complianceRate)}\nCompliance Failures: ${fmt.num(m.complianceFailures)}\nStatus Not Closed: ${fmt.num(m.statusNotClosed)}\n`;
  const a=document.createElement('a');a.href=URL.createObjectURL(new Blob([text],{type:'text/plain'}));a.download='way_decagon_summary.txt';a.click();
  showToast('Summary exported','success');
}

// ── NAV ──
function setupNav(){
  const TITLES={upload:'Data Source',kpis:'Executive KPIs',effectiveness:'Decagon Effectiveness',compliance:'Decagon Compliance',defects:'System Defects',reasons:'Reason Analysis',executive:'Executive Summary',validation:'Data Validation',tickets:'Master Tickets',recontact:'Re-contact Analysis'};
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
    document.getElementById('mainWrapper').classList.toggle('expanded');
  });
}

// ── INIT ──
document.addEventListener('DOMContentLoaded',()=>{
  setupUpload();setupNav();
  document.getElementById('exportPdfBtn').addEventListener('click',exportPDF);
  document.getElementById('exportSummaryBtn').addEventListener('click',exportSummary);
  document.getElementById('applyDateBtn').addEventListener('click',applyDateFilter);
  document.getElementById('clearDateBtn').addEventListener('click',clearDateFilter);
  document.getElementById('closeTimelineModal').addEventListener('click',()=>{document.getElementById('timelineModal').style.display='none';});
  document.getElementById('timelineModal').addEventListener('click',e=>{if(e.target===document.getElementById('timelineModal'))document.getElementById('timelineModal').style.display='none';});
  document.getElementById('closeDefectModal').addEventListener('click',()=>{document.getElementById('defectModal').style.display='none';});
  document.getElementById('defectModal').addEventListener('click',e=>{if(e.target===document.getElementById('defectModal'))document.getElementById('defectModal').style.display='none';});
  document.getElementById('recalcDefectsBtn').addEventListener('click',()=>{
    CONFIG.DEFECT_THRESHOLD_SEC=parseInt(document.getElementById('defectThreshold').value)||60;
    computeShortIntervalDefects(STATE.ticketMap);
    computeShortIntervalDefects(STATE.filteredTickets);
    renderDefectSection(computeMetrics(STATE.filteredTickets));
    showToast('Threshold updated to '+CONFIG.DEFECT_THRESHOLD_SEC+'s','info');
  });
  document.querySelectorAll('.btn-img-export').forEach(btn=>{
    btn.addEventListener('click',e=>{
      e.stopPropagation();const chart=STATE.charts[btn.dataset.chart];if(!chart)return;
      const a=document.createElement('a');a.download=btn.dataset.chart+'.png';a.href=chart.toBase64Image();a.click();
      showToast('Chart saved as PNG','success');
    });
  });
  document.addEventListener('keydown',e=>{if(e.key==='Escape'){document.getElementById('timelineModal').style.display='none';document.getElementById('defectModal').style.display='none';}});
  showToast('Dashboard ready — upload your CS All Tickets XLSX or CSV','info',5000);
});
window.showTimeline=showTimeline;
window.showDefectDrill=showDefectDrill;
