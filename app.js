'use strict';
/* WAY-DECAGON DASHBOARD v3.2 — All logic corrections applied */

const CONFIG = {
  AI_TYPE: 'AI-Agent Call',
  INTERNAL: ['TL Review','Manager Review','QC Audit','Select','User Reviews','BBB Reviews','App Feedback','Escalation Handled by TL','Escalation handled by Escalation Team','Escalation handled by Manager','Escalation handled by Ops Team'],
  CUSTOMER_FACING: ['AI-Agent Call','Call','Email','Chat','SMS'],
  EXCLUDED_REASONS: ['escalated','non escalated','not escalated','ai handled','ai-handled',''],
  DEFECT_THRESHOLD_SEC: 60
};

const STATE = {
  rawRows:[], ticketMap:new Map(), filteredTickets:new Map(),
  charts:{}, datatables:{}, colMap:{}, fcrBuilt:false, fcrDrillTable:null, fcrBuilt:false,
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

// ── SAFE DATE STRING PARSER (PST only, no timezone conversion) ──
function parseDateStr(d){
  if(!d)return null;
  const s=String(d).trim();
  const m=s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})/);
  if(m)return m[3]+'-'+String(m[1]).padStart(2,'0')+'-'+String(m[2]).padStart(2,'0');
  const m2=s.match(/^(\d{4})-(\d{2})-(\d{2})/);
  if(m2)return m2[1]+'-'+m2[2]+'-'+m2[3];
  return null;
}

// ── COLUMN MAP ──
const COL_C={ticketId:['Ticket ID','ticket_id'],ogi:['OGI','ogi'],interaction:['Interaction','Interaction Type'],intDate:['Interaction date','Interaction Date','Created Date'],intId:['Interaction ID'],reason:['Reason'],subReason:['Sub Reason','sub_reason'],action:['Action','Action Taken'],status:['Status'],userId:['UserID','User ID','user_id'],agent:['Agent Name'],vertical:['Vertical'],subVertical:['SubVertical','Sub Vertical'],ticketCreatedDate:['Ticket_created_date','Ticket Created Date']};
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
    const st=(tk.status||'').toLowerCase().trim();
    const act=(tk.actionTaken||'').toLowerCase().trim();
    tk.missingReason=!tk.reason;
    tk.missingSubReason=!tk.subReason;
    tk.wrongAction=!tk.actionTaken||act==='open';
    tk.openStatus=st==='open';
    tk.pendingStatus=['in progress','pending','on hold'].includes(st);
    tk.otherStatus=st!==''&&st!=='closed'&&!tk.openStatus&&!tk.pendingStatus;
    tk.statusNotClosed=st!=='closed';
    tk.missingAction=tk.wrongAction;
    tk.compliant=!tk.missingReason&&!tk.missingSubReason&&!tk.wrongAction&&!tk.openStatus&&!tk.pendingStatus&&!tk.otherStatus&&st==='closed';
  }else{tk.missingReason=tk.missingSubReason=tk.wrongAction=tk.missingAction=tk.openStatus=tk.pendingStatus=tk.otherStatus=tk.statusNotClosed=false;tk.compliant=false;}

  // FCR: Decagon-only + closed + has reason
  tk.fcrAchieved=false; // computed after computeShortIntervalDefects

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

  // Use parseDateStr to avoid timezone shift - all dates are PST
  const firstAIDateStr = firstAI ? parseDateStr(firstAI.createdDate) : null;
  if(firstAIDateStr){
    tk.aiInteractionDate = firstAIDateStr;
    tk.dateBucket = firstAIDateStr;
  } else if(tk.interactions[0]?.createdDate){
    const ds=parseDateStr(tk.interactions[0].createdDate);
    if(ds) tk.dateBucket=ds;
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
  const t_full=decOnly.filter(t=>t.compliant).length;
  const t_reason=decOnly.filter(t=>t.missingReason).length;
  const t_sub=decOnly.filter(t=>t.missingSubReason).length;
  const t_action=decOnly.filter(t=>t.wrongAction).length;
  const t_open=decOnly.filter(t=>t.openStatus).length;
  const t_pending=decOnly.filter(t=>t.pendingStatus).length;
  const t_other=decOnly.filter(t=>t.otherStatus).length;
  const missingReason=t_reason, missingSubReason=t_sub, missingAction=t_action;
  const statusNotClosed=decOnly.filter(t=>t.statusNotClosed).length;
  const pendingStatus=t_pending;
  const groupKey=t=>t.ogi&&t.ogi!=='UNKNOWN'?'O:'+t.ogi:(t.userId?'U:'+t.userId:'T:'+t.ticketId);
  const groups=new Map();
  decOnly.forEach(t=>{const k=groupKey(t);if(!groups.has(k))groups.set(k,{missingReason:false,missingSubReason:false,wrongAction:false,openStatus:false,pendingStatus:false,otherStatus:false,tickets:[]});const g=groups.get(k);g.tickets.push(t);if(t.missingReason)g.missingReason=true;if(t.missingSubReason)g.missingSubReason=true;if(t.wrongAction)g.wrongAction=true;if(t.openStatus)g.openStatus=true;if(t.pendingStatus)g.pendingStatus=true;if(t.otherStatus)g.otherStatus=true;});
  const uniqueGroups=groups.size;
  let g_full=0,g_reason=0,g_sub=0,g_action=0,g_open=0,g_pending=0,g_other=0;
  for(const g of groups.values()){
    if(g.missingReason)g_reason++;
    if(g.missingSubReason)g_sub++;
    if(g.wrongAction)g_action++;
    if(g.openStatus)g_open++;
    if(g.pendingStatus)g_pending++;
    if(g.otherStatus)g_other++;
    if(!g.missingReason&&!g.missingSubReason&&!g.wrongAction&&!g.openStatus&&!g.pendingStatus&&!g.otherStatus)g_full++;
  }
  const compliantCount=g_full;
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

  // Calculate filtered totals from rawRows filtered by date range
  const from=document.getElementById('globalDateFrom')?document.getElementById('globalDateFrom').value:'';
  const to=document.getElementById('globalDateTo')?document.getElementById('globalDateTo').value:'';
  const cm=STATE.colMap;
  let filteredTotalRecords=0,filteredTotalCallInts=0,filteredAIInts=0;
  console.log('[DEBUG] computeMetrics from='+from+' to='+to+' colMap.intDate='+cm.intDate+' rawRows='+STATE.rawRows.length);
  const CALL_TYPES=['Call','AI-Agent Call'];
  STATE.rawRows.forEach(row=>{
    const rawDate=cm.intDate?row[cm.intDate]:'';
    if(rawDate){
      const ds=parseDateStr(rawDate);
      if(!ds)return;
      if(from&&ds<from)return;
      if(to&&ds>to)return;
    }
    filteredTotalRecords++;
    const itype=cm.interaction?row[cm.interaction]:'';
    if(CALL_TYPES.includes(itype))filteredTotalCallInts++;
    if(itype==='AI-Agent Call')filteredAIInts++;
  });

  const _totalAI=filteredAIInts||STATE.totalAIInteractions;
  const avgAIPerTicket=decCount?(_totalAI/decCount):0;
  const multiAIDec=dec.filter(t=>t.aiInteractionCount>1).length;
  const repeatRate=decCount?(multiAIDec/decCount*100):0;
  const oneInX=repeatRate>0?Math.round(100/repeatRate):0;
  return{
    totalRecords:filteredTotalRecords||STATE.rawRows.length,
    totalCallInts:filteredTotalCallInts||STATE.totalCallInteractions,
    totalAIInts:_totalAI,
    avgAIPerTicket,repeatRate,oneInX,multiAIDec,
    totalTickets:all.length,
    decagonTickets:decCount,
    decagonOnlyCount:decOnlyCount,
    csAssistedCount,
    containedCount,containmentRate:pct(containedCount,decCount),
    fcrCount,fcrRate:pct(fcrCount,decCount),
    compliantCount,complianceRate:pct(g_full,uniqueGroups),
    complianceFailures:uniqueGroups-g_full,
    missingReason,missingSubReason,missingAction,statusNotClosed,pendingStatus,
    complianceGroups:uniqueGroups,g_full,g_reason,g_sub,g_action,g_open,g_pending,g_other,
    t_full,t_reason,t_sub,t_action,t_open,t_pending,t_other,
    ticketComplianceRate:pct(t_full,decOnlyCount),
    complianceGroupsMap:groups,
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
      if(!map.has(tid))map.set(tid,{ticketId:tid,ogi:getV(row,cm.ogi)||'UNKNOWN',userId:getV(row,cm.userId)||'',createdDate:getV(row,cm.ticketCreatedDate)||getV(row,cm.intDate),reason:getV(row,cm.reason),subReason:getV(row,cm.subReason),actionTaken:getV(row,cm.action),status:getV(row,cm.status),vertical:getV(row,cm.vertical),subVertical:getV(row,cm.subVertical),interactions:[]});
      const tk=map.get(tid);
      tk.interactions.push({interactionId:getV(row,cm.intId),type:getV(row,cm.interaction),createdDate:getV(row,cm.intDate),dateStr:parseDateStr(getV(row,cm.intDate)),parsedDate:parseDate(getV(row,cm.intDate)),reason:getV(row,cm.reason),subReason:getV(row,cm.subReason),actionTaken:getV(row,cm.action),agent:getV(row,cm.agent),status:getV(row,cm.status)});
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
          STATE.ticketMap.forEach(tk=>{tk.fcrAchieved=tk.isDecagonTicket&&!tk.csAssisted&&tk.aiInteractionCount===1&&!tk.shortIntervalFlag;});
          // fcrAchieved already set correctly above
          const m=computeMetrics(STATE.ticketMap);
          const allDates=[...map.values()].filter(t=>t.dateBucket).map(t=>t.dateBucket);
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
    {label:'Interactions by Decagon',mainVal:fmt.num(m.totalAIInts),subVal:`Avg ${m.avgAIPerTicket.toFixed(2)} AI interactions per ticket`,icon:'fa-comments',color:'purple',tip:`Total AI-Agent Call interaction records. ${fmt.num(m.multiAIDec)} tickets had >1 AI interaction (repeat rate: ${m.repeatRate.toFixed(1)}%)`,lvl:'Interaction'},
    {label:'Decagon FCR',mainVal:fmt.pct(m.fcrRate),subVal:fmt.num(m.fcrCount)+' calls',icon:'fa-bullseye',color:'green',tip:'FCR Met: single Decagon interaction with no further CS involvement or repeat contact. Base: all '+fmt.num(m.decagonTickets)+' Decagon tickets',lvl:'Ticket'},
    {label:'Decagon Containment Rate',mainVal:fmt.pct(m.containmentRate),subVal:fmt.num(m.containedCount)+' calls',icon:'fa-shield-halved',color:'green',tip:'Calls where no CS agent was involved after Decagon. 1,785 - 468 = 1,317',lvl:'Ticket'},
    {label:'CS Assisted',mainVal:fmt.pct(m.csAssistedCount/m.decagonTickets*100),subVal:fmt.num(m.csAssistedCount)+' calls',icon:'fa-person-walking-arrow-right',color:'amber',tip:'Calls where a human CS agent had to handle after Decagon',lvl:'Ticket',pctLarge:true},
    {label:'Handled by Decagon Only',mainVal:fmt.num(m.decagonOnlyCount),subVal:fmt.pct(pct(m.decagonOnlyCount,m.decagonTickets)),icon:'fa-circle-check',color:'green',tip:'Calls handled entirely by Decagon with no CS agent involvement',lvl:'Ticket'},
    {label:'Compliance Failures',mainVal:fmt.num(m.complianceFailures),subVal:fmt.num(m.complianceGroups)+' groups',icon:'fa-triangle-exclamation',color:'red',tip:'Unique orders/users with violations. Base: '+fmt.num(m.complianceGroups)+' unique orders/users (from '+fmt.num(m.decagonOnlyCount)+' Decagon-only calls)',lvl:'Order/User'},
    {label:'Compliance Rate',mainVal:fmt.pct(m.complianceRate),subVal:fmt.num(m.compliantCount)+' orders/users',icon:'fa-clipboard-check',color:'green',tip:'Unique orders/users with Reason + Sub Reason + valid Action + Status = Closed on ALL their Decagon-only tickets. Base: '+fmt.num(m.complianceGroups)+' unique orders/users',lvl:'Order/User'},
    {label:'Decagon Duplicate Ticket',mainVal:fmt.num(m.dupTicketCount),subVal:null,icon:'fa-copy',color:'red',tip:'Same OGI with multiple different Ticket IDs created at the same timestamp by Decagon',lvl:'Ticket'},
    {label:'Short Interval Interactions',mainVal:fmt.num(m.shortIntervalInts),subVal:null,icon:'fa-stopwatch',color:'amber',tip:'AI-Agent Call interactions on the same ticket within '+CONFIG.DEFECT_THRESHOLD_SEC+'s of each other — possible system retry',lvl:'Interaction'}
  ];

  document.getElementById('kpiGrid').innerHTML=kpis.map(k=>{
    const c=colorMap[k.color]||colorMap.cyan;
    const mainHtml=k.pctLarge
      ?`<div class="kpi-pct-large" style="color:${c.a}">${k.mainVal}</div>${k.subVal?`<div class="kpi-num-small">${k.subVal}</div>`:''}`
      :`<div class="kpi-val-large">${k.mainVal}</div>${k.subVal?`<div class="kpi-val-small">${k.subVal}</div>`:''}`;
    return`<div class="kpi-card" style="--ac:${c.a};--acd:${c.d}">
      <div class="kpi-card-top">
        <div class="kpi-icon"><i class="fa-solid ${k.icon}"></i></div>
        <div class="kpi-tip" title="${k.tip}" style="color:#cbd5e1;cursor:help"><i class="fa-solid fa-circle-info"></i></div>
      </div>
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
  const labels=[...buckets.keys()].map(d=>{const[y,mo,dy]=d.split('-');return new Date(y,mo-1,dy).toLocaleDateString('en-US',{month:'short',day:'numeric'});});
  const decCounts=[...buckets.values()].map(ts=>ts.filter(t=>t.isDecagonTicket).length);
  const intCounts=[...buckets.values()].map(ts=>ts.filter(t=>t.isDecagonTicket).length);
  // FCR: closed + no CS + has reason / decagonOnly per day
  const fcrRates=[...buckets.values()].map(ts=>{const d=ts.filter(t=>t.isDecagonTicket);return d.length?pct(d.filter(t=>t.fcrAchieved).length,d.length):0;});
  const csRates=[...buckets.values()].map(ts=>{const d=ts.filter(t=>t.isDecagonTicket);return d.length?pct(d.filter(t=>t.csAssisted).length,d.length):0;});
  const containRates=[...buckets.values()].map(ts=>{const d=ts.filter(t=>t.isDecagonTicket);return d.length?pct(d.filter(t=>t.decagonContained).length,d.length):0;});

  const base={responsive:true,maintainAspectRatio:true,interaction:{mode:'index',intersect:false},plugins:{legend:{labels:{color:text,font:{size:11}}},tooltip:{backgroundColor:'#fff',titleColor:'#0f172a',bodyColor:'#475569',borderColor:'#e2e8f0',borderWidth:1}},scales:{x:{ticks:{color:text,font:{size:10},maxRotation:45},grid:{color:grid}},y:{ticks:{color:text,font:{size:10}},grid:{color:grid},beginAtZero:true}}};

  const decOnlyCounts=[...buckets.values()].map(ts=>ts.filter(t=>t.isDecagonTicket&&t.decagonOnly).length);
  const csAssistedCounts=[...buckets.values()].map(ts=>ts.filter(t=>t.isDecagonTicket&&t.csAssisted).length);
  const stackLabelPlugin={id:'stackLabels',afterDatasetsDraw(chart){
    const ctx=chart.ctx;
    chart.data.datasets.forEach((ds,di)=>{
      chart.getDatasetMeta(di).data.forEach((bar,i)=>{
        const v=ds.data[i];if(!v||v<1)return;
        ctx.save();ctx.fillStyle='#fff';ctx.font='bold 9px sans-serif';ctx.textAlign='center';ctx.textBaseline='middle';
        const y=bar.y+(bar.base-bar.y)/2;
        if(bar.base-bar.y>14)ctx.fillText(v,bar.x,y);
        ctx.restore();
      });
    });
  }};
  dChart('decagonTicketsTrend');
  STATE.charts.decagonTicketsTrend=new Chart(document.getElementById('decagonTicketsTrend'),{type:'bar',data:{labels,datasets:[
    {label:'Decagon Only',data:decOnlyCounts,backgroundColor:'rgba(2,132,199,0.6)'},
    {label:'CS Assisted',data:csAssistedCounts,backgroundColor:'rgba(239,68,68,0.7)'}
  ]},options:{...base,plugins:{...base.plugins,legend:{display:true,position:'top',labels:{color:text,font:{size:11}}}},scales:{...base.scales,x:{...base.scales?.x,stacked:true,ticks:{color:text,font:{size:10},maxRotation:45},grid:{color:grid}},y:{...base.scales?.y,stacked:true,ticks:{color:text,font:{size:10}},grid:{color:grid},beginAtZero:true}}},plugins:[stackLabelPlugin]});

  const csIntCounts=[...buckets.values()].map(ts=>ts.filter(t=>t.isDecagonTicket).reduce((s,t)=>s+t.humanInteractionCount,0));
  dChart('decagonIntsTrend');
  const stackLabelPlugin2={id:'stackLabels2',afterDatasetsDraw(chart){
    const ctx=chart.ctx;
    chart.data.datasets.forEach((ds,di)=>{
      chart.getDatasetMeta(di).data.forEach((bar,i)=>{
        const v=ds.data[i];
        if(!v||v<1)return;
        ctx.save();ctx.fillStyle='#fff';ctx.font='bold 9px sans-serif';ctx.textAlign='center';ctx.textBaseline='middle';
        const y=bar.y+(bar.base-bar.y)/2;
        if(bar.base-bar.y>14)ctx.fillText(v,bar.x,y);
        ctx.restore();
      });
    });
  }};
  STATE.charts.decagonIntsTrend=new Chart(document.getElementById('decagonIntsTrend'),{type:'bar',data:{labels,datasets:[
    {label:'AI Interactions',data:intCounts,backgroundColor:'rgba(124,58,237,0.6)'},
    {label:'CS Interactions',data:csIntCounts,backgroundColor:'rgba(239,68,68,0.7)'}
  ]},options:{...base,plugins:{...base.plugins,legend:{display:true,position:'top',labels:{color:text,font:{size:11}}}},scales:{...base.scales,x:{...base.scales?.x,stacked:true,ticks:{color:text,font:{size:10},maxRotation:45},grid:{color:grid}},y:{...base.scales?.y,stacked:true,ticks:{color:text,font:{size:10}},grid:{color:grid},beginAtZero:true}}},plugins:[stackLabelPlugin2]});

  dChart('fcrTrend');
  STATE.charts.fcrTrend=new Chart(document.getElementById('fcrTrend'),{type:'line',data:{labels,datasets:[{label:'Decagon FCR %',data:fcrRates,borderColor:'#059669',backgroundColor:'rgba(5,150,105,0.1)',fill:true,tension:0.4,pointRadius:3}]},options:{...base,scales:{x:{ticks:{color:text,font:{size:10},maxRotation:45},grid:{color:grid}},y:{ticks:{color:text,font:{size:10},callback:v=>v+'%'},grid:{color:grid},beginAtZero:true,max:100}}}});

  dChart('csAssistedTrend');
  STATE.charts.csAssistedTrend=new Chart(document.getElementById('csAssistedTrend'),{type:'line',data:{labels,datasets:[{label:'CS Assisted %',data:csRates,borderColor:'#d97706',backgroundColor:'rgba(217,119,6,0.08)',fill:true,tension:0.4,pointRadius:3},{label:'Containment %',data:containRates,borderColor:'#059669',fill:false,tension:0.4,pointRadius:3,borderDash:[5,4]}]},options:{...base,scales:{x:{ticks:{color:text,font:{size:10},maxRotation:45},grid:{color:grid}},y:{ticks:{color:text,font:{size:10},callback:v=>v+'%'},grid:{color:grid},beginAtZero:true}}}});
}

// ── COMPLIANCE ──
function renderComplianceSection(m){
  document.getElementById('gaugeCompPct').textContent=fmt.pct(m.complianceRate);
  document.getElementById('cv-full').textContent=fmt.num(m.g_full)+' orders/users';
  document.getElementById('cv-reason').textContent=fmt.num(m.g_reason);
  document.getElementById('cv-sub').textContent=fmt.num(m.g_sub);
  document.getElementById('cv-action').textContent=fmt.num(m.g_action);
  document.getElementById('cv-open').textContent=fmt.num(m.g_open);
  document.getElementById('cv-pending').textContent=fmt.num(m.g_pending);
  const cvOther=document.getElementById('cv-other'),cvOtherRow=document.getElementById('cv-other-row');
  if(cvOther){cvOther.textContent=fmt.num(m.g_other);cvOtherRow.style.display=m.g_other>0?'flex':'none';}
  const setTv=(id,val)=>{const el=document.getElementById(id);if(el)el.textContent=fmt.num(val);};
  setTv('tv-full',m.t_full); setTv('tv-reason',m.t_reason); setTv('tv-sub',m.t_sub);
  setTv('tv-action',m.t_action); setTv('tv-open',m.t_open); setTv('tv-pending',m.t_pending);
  const tvOtherRow=document.getElementById('tv-other-row');
  setTv('tv-other',m.t_other); if(tvOtherRow)tvOtherRow.style.display=m.t_other>0?'flex':'none';
  const tvFullEl=document.getElementById('tv-full');if(tvFullEl)tvFullEl.textContent=fmt.num(m.t_full)+' tickets';
  const hdr=document.querySelector('#tab-compliance .panel-header p');
  if(hdr)hdr.textContent='Group-level view · '+fmt.num(m.complianceGroups)+' unique orders/users from '+fmt.num(m.decagonOnlyCount)+' Decagon-only calls · Group by OGI → User ID → Ticket ID · Strict OR on violations';
  renderGauge('complianceGauge',m.complianceRate);
  const{text,grid}=getCC();

  dChart('compliancePie');
  STATE.charts.compliancePie=new Chart(document.getElementById('compliancePie'),{
    type:'doughnut',
    data:{labels:['Fully Compliant','Missing Reason','Missing Sub Reason','Wrong Action','Open Status','Pending Status','Other Status'].slice(0,m.g_other>0?7:6),datasets:[{data:[m.g_full,m.g_reason,m.g_sub,m.g_action,m.g_open,m.g_pending,m.g_other].slice(0,m.g_other>0?7:6),backgroundColor:['#10b981','#ef4444','#f59e0b','#f97316','#8b5cf6','#6366f1','#64748b'],borderColor:'#fff',borderWidth:2}]},
    options:{responsive:true,maintainAspectRatio:true,plugins:{legend:{position:'bottom',labels:{color:text,font:{size:11},padding:8}}}}
  });

  const buckets=getDateBuckets(STATE.filteredTickets);
  const labels=[...buckets.keys()].map(d=>{const[y,mo,dy]=d.split('-');return new Date(y,mo-1,dy).toLocaleDateString('en-US',{month:'short',day:'numeric'});});
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
        if(drill==='wrongAction')return t.wrongAction;
        if(drill==='openStatus')return t.openStatus;
        if(drill==='pendingStatus')return t.pendingStatus;
        if(drill==='otherStatus')return t.otherStatus;
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
  const trueTotalCalls=m.totalCallInts+m.decagonOnlyCount; const decagonShareOfCalls=pct(m.decagonOnlyCount,trueTotalCalls);
  const callsRoutedToDecagon=m.decagonTickets; // 1785
  const callsHandledAlone=m.decagonOnlyCount; // 1317
  const pctHandledAlone=pct(callsHandledAlone,callsRoutedToDecagon);

  // Key observations
  const obs=[];
  obs.push(`Decagon handled <strong>${fmt.num(m.decagonTickets)} calls</strong> out of <strong>${fmt.num(trueTotalCalls)} total calls</strong> — representing <strong>${fmt.pct(decagonShareOfCalls)}</strong> of all voice interactions in the CRM.`);
  obs.push(`Of ${fmt.num(m.decagonTickets)} calls routed to Decagon, <strong>${fmt.num(callsHandledAlone)} (${fmt.pct(pctHandledAlone)})</strong> were handled by Decagon alone without CS involvement.`);
  obs.push(`<strong>${fmt.num(m.fcrCount)} tickets (${fmt.pct(m.fcrRate)})</strong> met FCR — single Decagon interaction with no further contact from the customer. Remaining <strong>${fmt.num(m.decagonTickets-m.fcrCount)}</strong> tickets had CS involvement, multiple AI interactions, or repeat contacts.`);
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
      <div class="ceo-meta-item"><div class="ceo-meta-label">Total Calls (Human + Decagon)</div><div class="ceo-meta-val">${fmt.num(trueTotalCalls)}</div></div>
      <div class="ceo-meta-item"><div class="ceo-meta-label">Calls Routed to Decagon</div><div class="ceo-meta-val">${fmt.num(callsRoutedToDecagon)}</div></div>
      <div class="ceo-meta-item"><div class="ceo-meta-label">Calls Handled by Decagon Alone</div><div class="ceo-meta-val">${fmt.num(callsHandledAlone)}</div></div>
    </div>
    <div class="ceo-kpi-row">
      <div class="ceo-kpi-item"><div class="ceo-kpi-label">Escalated to CS</div><div class="ceo-kpi-val" style="color:#d97706">${fmt.num(m.csAssistedCount)} <span style="font-size:12px">(${fmt.pct(m.csAssistedCount/m.decagonTickets*100)})</span></div></div>
      <div class="ceo-kpi-item"><div class="ceo-kpi-label">Decagon FCR</div><div class="ceo-kpi-val" style="color:#dc2626">${fmt.pct(m.fcrRate)}</div></div>
      <div class="ceo-kpi-item"><div class="ceo-kpi-label">Containment Rate</div><div class="ceo-kpi-val" style="color:#059669">${fmt.pct(m.containmentRate)}</div></div>
      <div class="ceo-kpi-item"><div class="ceo-kpi-label">Compliance Rate <span style="font-size:10px;font-weight:400;color:#64748b">(Order/User)</span></div><div class="ceo-kpi-val" style="color:#dc2626">${fmt.pct(m.complianceRate)}</div></div>
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
  STATE.fcrBuilt=false;
  STATE.ticketMap.forEach((tk,id)=>{
    // Use YYYY-MM-DD string comparison to avoid timezone issues
    const filterDate = tk.dateBucket || tk.aiInteractionDate;
    if(!filterDate)return;
    if(from&&filterDate<from)return;
    if(to&&filterDate>to)return;
    STATE.filteredTickets.set(id,tk);
  });
  renderDashboard();showToast('Filter applied — '+fmt.num(STATE.filteredTickets.size)+' tickets','info');
}

function clearDateFilter(){
  STATE.filteredTickets=new Map(STATE.ticketMap);
  const dates=[...STATE.ticketMap.values()].filter(t=>t.dateBucket).map(t=>t.dateBucket);
  if(dates.length){dates.sort();document.getElementById('globalDateFrom').value=dates[0];document.getElementById('globalDateTo').value=dates[dates.length-1];}
  renderDashboard();showToast('Filter cleared','info');
}

// ── EXPORT ──
function buildTrendsTab() {
  const el = document.getElementById('trendsContent');
  if (!el) return;
  const allTickets = [...STATE.ticketMap.values()];

  function parseDStr(s) {
    if (!s) return null;
    const m = s.match(/(\d+)\/(\d+)\/(\d+)\s+(\d+):(\d+):(\d+)\s+(AM|PM)/i);
    if (!m) return null;
    let [,mo,dy,yr,hr,mn,sc,ap] = m;
    hr = parseInt(hr);
    if (ap.toUpperCase()==='PM'&&hr!==12) hr+=12;
    if (ap.toUpperCase()==='AM'&&hr===12) hr=0;
    return new Date(yr, mo-1, dy, hr, mn, sc);
  }

  function getTicketDate(tk) {
    const ai = tk.interactions.filter(i => i.type==='AI-Agent Call' && i.dateStr);
    if (!ai.length) return null;
    // Use earliest AI interaction date
    return ai.reduce((a,b) => a.dateStr < b.dateStr ? a : b).dateStr;
  }

  function dayKey(ds) { return ds || null; }
  function weekKey(ds) {
    if (!ds) return null;
    const [y,m,d] = ds.split('-').map(Number);
    const t = new Date(y, m-1, d);
    const day = t.getDay();
    const diff = day === 0 ? -6 : 1 - day;
    t.setDate(t.getDate() + diff);
    return t.getFullYear()+'-'+String(t.getMonth()+1).padStart(2,'0')+'-'+String(t.getDate()).padStart(2,'0');
  }
  function monthKey(ds) { return ds ? ds.slice(0,7) : null; }
  function yearKey(ds) { return ds ? ds.slice(0,4) : null; }

  function keyLabel(k, mode) {
    if (mode==='daily') { const [y,m,d]=k.split('-'); return new Date(y,m-1,d).toLocaleDateString('en-GB',{day:'2-digit',month:'short'}); }
    if (mode==='weekly') {
      const [y,m,d]=k.split('-').map(Number);
      const mon=new Date(y,m-1,d);
      const sun=new Date(y,m-1,d+6);
      return mon.toLocaleDateString('en-GB',{day:'2-digit',month:'short'})+' – '+sun.toLocaleDateString('en-GB',{day:'2-digit',month:'short'});
    }
    if (mode==='monthly') { const [y,m]=k.split('-'); return new Date(y,m-1,1).toLocaleDateString('en-GB',{month:'short',year:'numeric'}); }
    return k;
  }

  function buildBuckets(keyFn) {
    const B = {};
    const ensure = k => { if (!B[k]) B[k]={totalCalls:0,csCalls:0,decCalls:0,decTickets:0,fullHandled:0,escalated:0,fcrMet:0,contained:0,compliant:0,decOnly:0,repCalls:0,repCustomers:0}; };
    for (const tk of allTickets) {
      for (const i of tk.interactions) {
        if (i.type!=='Call'&&i.type!=='AI-Agent Call') continue;
        const k=keyFn(i.dateStr);
        if (!k) continue;
        ensure(k);
        if (i.type==='Call') { B[k].totalCalls++; B[k].csCalls++; }
      }
    }
    const repCust = {};
    for (const tk of allTickets) {
      if (!tk.isDecagonTicket) continue;
      const k=keyFn(getTicketDate(tk));
      if (!k) continue;
      ensure(k);
      B[k].decTickets++;
      if (tk.decagonOnly){B[k].fullHandled++;B[k].decOnly++;B[k].decCalls++;B[k].totalCalls++;}
      if (tk.csAssisted) B[k].escalated++;
      if (tk.fcrAchieved) B[k].fcrMet++;
      B[k].totalDecTickets = (B[k].totalDecTickets||0) + 1;
      if (tk.decagonContained) B[k].contained++;
      if (tk.compliant) B[k].compliant++;
      if (tk.aiInteractionCount>1){
        B[k].repCalls+=tk.aiInteractionCount-1;
        if (!repCust[k]) repCust[k]=new Set();
        if (tk.ogi) repCust[k].add(tk.ogi); else repCust[k].add(tk.ticketId);
      }
    }
    for (const [k,s] of Object.entries(repCust)) { if (B[k]) B[k].repCustomers=s.size; }
    return B;
  }

  let mode='weekly', offset=0, dailyB, weeklyB, monthlyB, yearlyB;

  function getBuckets() {
    if (mode==='daily'){if(!dailyB)dailyB=buildBuckets(dayKey);return dailyB;}
    if (mode==='weekly'){if(!weeklyB)weeklyB=buildBuckets(weekKey);return weeklyB;}
    if (mode==='monthly'){if(!monthlyB)monthlyB=buildBuckets(monthKey);return monthlyB;}
    if(!yearlyB)yearlyB=buildBuckets(yearKey);return yearlyB;
  }

  const fmt = n => (n==null||isNaN(n))?'—':Math.round(n).toLocaleString();
  const pct = (a,b) => b?(a/b*100).toFixed(1)+'%':'—';
  function delta(c,p) {
    if (p==null||c==null) return '<span style="font-size:11px;color:var(--color-text-secondary)">—</span>';
    const d=c-p; if (!d) return '<span style="font-size:11px;color:var(--color-text-secondary)">—</span>';
    return d>0?`<span style="font-size:11px;font-weight:500;color:#3B6D11">▲ ${fmt(d)}</span>`:`<span style="font-size:11px;font-weight:500;color:#A32D2D">▼ ${fmt(Math.abs(d))}</span>`;
  }

  const ROWS=[
    {section:'Call volume'},
    {key:'totalCalls',label:'Total calls handled',color:'#185FA5'},
    {key:'csCalls',label:'Calls handled by CS',color:'#3B6D11',pctOf:'totalCalls'},
    {key:'decCalls',label:'Calls handled by Decagon',color:'#534AB7',pctOf:'totalCalls'},
    {key:'repCalls',label:'Repeat Decagon contacts',color:'#854F0B',extra:'repCustomers'},
    {section:'Decagon performance'},
    {key:'fullHandled',label:'Fully handled by Decagon',color:'#639922',pctOf:'decTickets'},
    {key:'escalated',label:'Escalated to CS by Decagon',color:'#993C1D',pctOf:'decTickets'},
    {key:'fcrMet',label:'Decagon FCR',color:'#1D9E75',pctOf:'decTickets',isPct:true},
    {key:'contained',label:'Decagon containment rate',color:'#0F6E56',pctOf:'decTickets',isPct:true},
    {key:'compliant',label:'Decagon compliance rate',color:'#7F77DD',pctOf:'decOnly',isPct:true},
  ];

  function render() {
    const B=getBuckets();
    const keys=Object.keys(B).sort();
    const ps=mode==='daily'?7:mode==='weekly'?4:mode==='monthly'?12:keys.length;
    const maxOff=Math.max(0,Math.floor((keys.length-ps)/ps));
    // Pin baseline: skip buckets before Jun 1
    const junStart=mode==='monthly'?'2026-06':mode==='yearly'?'2026':'2026-06-01';
    const firstJunIdx=keys.findIndex(k=>k>=junStart);
    const baseIdx=firstJunIdx>=0?firstJunIdx:0;
    if(offset===0)offset=0; // keep user nav relative to Jun start
    const si=Math.max(baseIdx,Math.min(keys.length-ps,baseIdx+offset*ps));
    const cols=keys.slice(si,si+ps);
    if(!cols.length){el.innerHTML='<p style="padding:2rem;color:var(--color-text-secondary)">No data available</p>';return;}

    const allKeys=Object.keys(B);
    const totP=allKeys.reduce((s,k)=>s+(B[k].totalCalls||0),0);
    const csP=allKeys.reduce((s,k)=>s+(B[k].csCalls||0),0);
    const decP=allKeys.reduce((s,k)=>s+(B[k].decCalls||0),0);
    const pLabel=keyLabel(cols[0],mode)+' → '+keyLabel(cols[cols.length-1],mode);

    const mBtn=m=>`<button onclick="window._tMode('${m}')" style="padding:6px 14px;font-size:12px;font-weight:500;border:${mode===m?'0.5px solid var(--color-border-secondary)':'none'};background:${mode===m?'var(--color-background-primary)':'transparent'};color:${mode===m?'var(--color-text-primary)':'var(--color-text-secondary)'};border-radius:6px;cursor:pointer">${m.charAt(0).toUpperCase()+m.slice(1)}</button>`;

    // Build sparkline data from ALL buckets
    const sparkKeys = Object.keys(B).sort();
    const sparkTotal = sparkKeys.map(k=>B[k].totalCalls||0);
    const sparkCS = sparkKeys.map(k=>B[k].csCalls||0);
    const sparkDec = sparkKeys.map(k=>B[k].decCalls||0);

    // SVG sparkline helper
    function sparkline(data, color) {
      if (!data.length) return '';
      const w=120,h=40,pad=4;
      const max=Math.max(...data)||1, min=Math.min(...data);
      const range=max-min||1;
      const pts=data.map((v,i)=>{
        const x=pad+(i/(data.length-1||1))*(w-pad*2);
        const y=h-pad-(v-min)/range*(h-pad*2);
        return x+','+y;
      });
      return `<svg width="${w}" height="${h}" style="overflow:visible"><polyline points="${pts.join(' ')}" fill="none" stroke="${color}" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"/><circle cx="${pts[pts.length-1].split(',')[0]}" cy="${pts[pts.length-1].split(',')[1]}" r="3" fill="${color}"/></svg>`;
    }

    // Period-over-period delta for KPI tiles (compare current visible period vs previous same-length period)
    const prevKeys = sparkKeys.slice(Math.max(0,sparkKeys.indexOf(cols[0])-cols.length), sparkKeys.indexOf(cols[0]));
    const prevTotP = prevKeys.reduce((s,k)=>s+(B[k]?.totalCalls||0),0);
    const prevCsP = prevKeys.reduce((s,k)=>s+(B[k]?.csCalls||0),0);
    const prevDecP = prevKeys.reduce((s,k)=>s+(B[k]?.decCalls||0),0);

    function kpiDelta(curr, prev) {
      if (!prev) return '';
      const pct2 = ((curr-prev)/prev*100);
      const col = pct2>=0?'#059669':'#dc2626';
      const arr = pct2>=0?'▲':'▼';
      return `<span style="font-size:12px;font-weight:600;color:${col}">${arr} ${Math.abs(pct2).toFixed(1)}%</span> <span style="font-size:11px;color:#94a3b8">vs prior period</span>`;
    }

    let h=`<div style="padding:1.5rem;background:#f8fffe;min-height:100%">
    <div style="display:flex;align-items:center;justify-content:space-between;flex-wrap:wrap;gap:8px;margin-bottom:20px">
      <div style="display:flex;gap:4px;background:#fff;padding:4px;border-radius:8px;border:1px solid #cce8e8;box-shadow:0 1px 4px rgba(13,148,136,0.08)">${['daily','weekly','monthly','yearly'].map(mBtn).join('')}</div>
      <div style="display:flex;align-items:center;gap:10px">
        <button onclick="window._tNav(-1)" style="width:32px;height:32px;border:1px solid #cce8e8;background:#fff;border-radius:8px;cursor:pointer;font-size:16px;color:#0d9488;font-weight:600">←</button>
        <span style="font-size:13px;font-weight:600;color:#0f172a;min-width:220px;text-align:center">${pLabel}</span>
        <button onclick="window._tNav(1)" style="width:32px;height:32px;border:1px solid #cce8e8;background:#fff;border-radius:8px;cursor:pointer;font-size:16px;color:#0d9488;font-weight:600">→</button>
      </div>
      <span style="font-size:11px;color:#94a3b8;font-style:italic">Comparing to prior period</span>
    </div>
    <div style="display:grid;grid-template-columns:repeat(3,1fr);gap:16px;margin-bottom:24px">
      <div style="background:#fff;border-radius:12px;padding:20px;border:1px solid #cce8e8;box-shadow:0 2px 8px rgba(13,148,136,0.08);display:flex;flex-direction:column;gap:4px">
        <div style="display:flex;justify-content:space-between;align-items:flex-start">
          <div>
            <div style="font-size:10px;font-weight:700;color:#0d9488;text-transform:uppercase;letter-spacing:0.06em;margin-bottom:6px">Total calls handled</div>
            <div style="font-size:32px;font-weight:800;color:#0f172a;line-height:1;font-family:monospace">${fmt(totP)}</div>
            <div style="font-size:12px;color:#64748b;margin-top:4px">Full dataset total</div>
          </div>
          <div style="opacity:0.7">${sparkline(sparkTotal,'#0d9488')}</div>
        </div>
        <div style="margin-top:8px">${kpiDelta(totP,prevTotP)}</div>
      </div>
      <div style="background:#fff;border-radius:12px;padding:20px;border:1px solid #cce8e8;box-shadow:0 2px 8px rgba(13,148,136,0.08);display:flex;flex-direction:column;gap:4px">
        <div style="display:flex;justify-content:space-between;align-items:flex-start">
          <div>
            <div style="font-size:10px;font-weight:700;color:#059669;text-transform:uppercase;letter-spacing:0.06em;margin-bottom:6px">Calls handled by CS</div>
            <div style="font-size:32px;font-weight:800;color:#059669;line-height:1;font-family:monospace">${fmt(csP)}</div>
            <div style="font-size:12px;color:#64748b;margin-top:4px">${pct(csP,totP)} of total calls</div>
          </div>
          <div style="opacity:0.7">${sparkline(sparkCS,'#059669')}</div>
        </div>
        <div style="margin-top:8px">${kpiDelta(csP,prevCsP)}</div>
      </div>
      <div style="background:#fff;border-radius:12px;padding:20px;border:1px solid #cce8e8;box-shadow:0 2px 8px rgba(13,148,136,0.08);display:flex;flex-direction:column;gap:4px">
        <div style="display:flex;justify-content:space-between;align-items:flex-start">
          <div>
            <div style="font-size:10px;font-weight:700;color:#7c3aed;text-transform:uppercase;letter-spacing:0.06em;margin-bottom:6px">Calls handled by Decagon</div>
            <div style="font-size:32px;font-weight:800;color:#7c3aed;line-height:1;font-family:monospace">${fmt(decP)}</div>
            <div style="font-size:12px;color:#64748b;margin-top:4px">${pct(decP,totP)} of total calls</div>
          </div>
          <div style="opacity:0.7">${sparkline(sparkDec,'#7c3aed')}</div>
        </div>
        <div style="margin-top:8px">${kpiDelta(decP,prevDecP)}</div>
      </div>
    </div>
`;

    const ICONS2={totalCalls:'fa-phone',csCalls:'fa-headset',decCalls:'fa-robot',repCalls:'fa-rotate',fullHandled:'fa-circle-check',escalated:'fa-person-walking-arrow-right',fcrMet:'fa-bullseye',contained:'fa-shield-halved',compliant:'fa-clipboard-check'};
    var tbl='<div style="overflow-x:auto;border-radius:12px;border:1px solid #cce8e8;box-shadow:0 2px 8px rgba(13,148,136,0.08);background:#fff">';
    tbl+='<table style="width:100%;border-collapse:collapse;font-size:13px;table-layout:fixed">';
    tbl+='<thead><tr style="background:#f0fafa;border-bottom:2px solid #cce8e8">';
    tbl+='<th style="padding:14px 16px;text-align:left;font-weight:700;font-size:11px;color:#475569;border-right:1px solid #cce8e8;width:230px;text-transform:uppercase;letter-spacing:0.05em">Metric</th>';
    cols.forEach(function(k,i){var lbl=keyLabel(k,mode);var sub=i===0?'<div style="font-size:10px;font-weight:600;color:#0d9488;margin-top:2px">Baseline</div>':'<div style="font-size:10px;font-weight:500;color:#94a3b8;margin-top:2px">vs prev</div>';tbl+='<th style="padding:10px 16px;text-align:center;font-weight:600;font-size:12px;color:#0f172a;border-right:1px solid #cce8e8"><div style="color:#0d9488;font-weight:700">'+lbl+'</div>'+sub+'</th>';});
    tbl+='</tr></thead><tbody>';
    for(var ri=0;ri<ROWS.length;ri++){
      var row=ROWS[ri];
      if(row.section){tbl+='<tr><td colspan="'+(cols.length+1)+'" style="padding:8px 16px;font-size:10px;font-weight:700;color:#0d9488;background:#f0fafa;border-bottom:1px solid #cce8e8;text-transform:uppercase;letter-spacing:0.08em;border-left:4px solid #0d9488">'+row.section+'</td></tr>';continue;}
      var icon=ICONS2[row.key]||'fa-chart-bar';
      var pof=row.pctOf==='totalCalls'?'total calls':row.pctOf==='decTickets'?'Decagon tickets':row.pctOf==='decOnly'?'Decagon only':'';
      tbl+='<tr style="border-bottom:1px solid #e6f5f5"><td style="padding:12px 16px;background:#fafffe;border-right:1px solid #cce8e8;border-left:3px solid '+row.color+'">';
      tbl+='<div style="display:flex;align-items:center;gap:8px"><span style="width:26px;height:26px;border-radius:6px;background:'+row.color+'22;display:inline-flex;align-items:center;justify-content:center;flex-shrink:0"><i class="fa-solid '+icon+'" style="font-size:11px;color:'+row.color+'"></i></span>';
      tbl+='<div><div style="font-size:12px;font-weight:600;color:#1e293b">'+row.label+'</div>';
      if(pof)tbl+='<div style="font-size:10px;color:#94a3b8">% of '+pof+'</div>';
      tbl+='</div></div></td>';
      cols.forEach(function(k,i){
        var b=B[k];var prev=i>0?B[cols[i-1]]:null;var val=b[row.key]||0;var isB=i===0;
        var disp='';var cL='';var sL='';
        if(row.isPct){
          var base=b[row.pctOf]||0;var cp=base?val/base*100:null;
          disp=cp!=null?cp.toFixed(1)+'%':'--';
          if(prev&&!isB){var pb=prev[row.pctOf]||0;var pp=pb?(prev[row.key]||0)/pb*100:null;if(pp!=null&&cp!=null){var d=cp-pp;if(Math.abs(d)>=0.01){var c2=d>0?'#059669':'#dc2626';cL='<div style="font-size:11px;font-weight:600;color:'+c2+';margin-top:3px">'+(d>0?'&#9650;':'&#9660;')+' '+Math.abs(d).toFixed(1)+'pp</div>';}}}
        } else {
          disp=fmt(val);
          if(prev&&!isB){var pv=prev[row.key]||0;var d=val-pv;var pc=pv?d/pv*100:null;if(d!==0&&pc!=null){var c2=d>0?'#059669':'#dc2626';cL='<div style="font-size:11px;font-weight:600;color:'+c2+';margin-top:3px">'+(d>0?'&#9650;':'&#9660;')+' '+Math.abs(pc).toFixed(1)+'%</div>';}}
          if(row.pctOf){var b2=b[row.pctOf]||0;if(b2)sL='<div style="font-size:11px;color:#64748b;margin-top:1px">'+pct(val,b2)+'</div>';}
          if(row.extra&&b[row.extra])sL+='<div style="font-size:11px;color:#64748b">'+fmt(b[row.extra])+' customers</div>';
        }
        tbl+='<td style="padding:12px 16px;text-align:center;border-right:1px solid #e6f5f5;background:'+(isB?'#f0fafa':'#fff')+'">';
        tbl+='<div style="font-size:16px;font-weight:700;color:'+(isB?'#0d9488':'#0f172a')+';font-family:monospace">'+disp+'</div>';
        tbl+=sL+cL+'</td>';
      });
      tbl+='</tr>';
    }
    tbl+='</tbody></table></div>';
    tbl+='<div style="margin-top:12px;display:flex;gap:20px;font-size:11px;color:#64748b;align-items:center">';
    tbl+='<span><span style="color:#059669;font-weight:600">&#9650;</span> Increase vs prior period</span>';
    tbl+='<span><span style="color:#dc2626;font-weight:600">&#9660;</span> Decrease vs prior period</span>';
    tbl+='<span style="color:#0d9488">pp = percentage points</span>';
    tbl+='<span style="margin-left:auto;font-style:italic">Unaffected by global date filter</span>';
    tbl+='</div></div>';
    var insFCRArr=cols.map(k=>B[k]&&B[k].decTickets?B[k].fcrMet/B[k].decTickets*100:null).filter(v=>v!=null);
    var insAvgFCR=insFCRArr.length?insFCRArr.reduce((a,b)=>a+b,0)/insFCRArr.length:0;
    var insContArr=cols.map(k=>B[k]&&B[k].decTickets?B[k].contained/B[k].decTickets*100:null).filter(v=>v!=null);
    var insAvgCont=insContArr.length?insContArr.reduce((a,b)=>a+b,0)/insContArr.length:0;
    var insCompArr=cols.map(k=>B[k]&&B[k].decOnly?B[k].compliant/B[k].decOnly*100:null).filter(v=>v!=null);
    var insAvgComp=insCompArr.length?insCompArr.reduce((a,b)=>a+b,0)/insCompArr.length:0;
    var insMid=Math.floor(cols.length/2);
    var insFCR1=cols.slice(0,insMid).map(k=>B[k]&&B[k].decTickets?B[k].fcrMet/B[k].decTickets*100:null).filter(v=>v!=null);
    var insFCR2=cols.slice(insMid).map(k=>B[k]&&B[k].decTickets?B[k].fcrMet/B[k].decTickets*100:null).filter(v=>v!=null);
    var insF1=insFCR1.length?insFCR1.reduce((a,b)=>a+b,0)/insFCR1.length:0;
    var insF2=insFCR2.length?insFCR2.reduce((a,b)=>a+b,0)/insFCR2.length:0;
    var insFCRTrend=insF2-insF1;
    var insFCRByPeriod=cols.map(k=>({k,rate:B[k]&&B[k].decTickets?B[k].fcrMet/B[k].decTickets*100:0})).filter(v=>v.rate>0);
    insFCRByPeriod.sort((a,b)=>b.rate-a.rate);
    var insBestFCR=insFCRByPeriod[0];var insWorstFCR=insFCRByPeriod[insFCRByPeriod.length-1];
    var insEscSpike=cols.filter(k=>B[k]&&B[k].decTickets&&(B[k].decCalls-B[k].fullHandled)/B[k].decTickets>0.5).length;
    var insCompGap=100-insAvgComp;
    var insRepRate=cols.reduce((s,k)=>s+(B[k]?B[k].repCalls||0:0),0);
    var insTotalDec=cols.reduce((s,k)=>s+(B[k]?B[k].decCalls||0:0),0);
    var insRepPct=insTotalDec?insRepRate/insTotalDec*100:0;
    function insIcon(type){return{good:'<span style="color:#059669;font-weight:700">✓</span>',warn:'<span style="color:#d97706;font-weight:700">⚠</span>',bad:'<span style="color:#dc2626;font-weight:700">✗</span>',info:'<span style="color:#0d9488;font-weight:700">→</span>'}[type]||'';}
    var panels=[];
    if(insFCRTrend>1)panels.push({type:'good',title:'FCR Improving',body:'FCR up <strong>'+Math.abs(insFCRTrend).toFixed(1)+'pp</strong> in the latter half of this period.'});
    else if(insFCRTrend<-1)panels.push({type:'warn',title:'FCR Declining',body:'FCR dropped <strong>'+Math.abs(insFCRTrend).toFixed(1)+'pp</strong> — review re-contact patterns.'});
    else panels.push({type:'info',title:'FCR Stable',body:'FCR steady at <strong>~'+insAvgFCR.toFixed(1)+'%</strong> with no significant drift.'});
    if(insBestFCR)panels.push({type:'good',title:'Best Period',body:'Highest FCR: <strong>'+insBestFCR.rate.toFixed(1)+'%</strong> — '+keyLabel(insBestFCR.k,mode)+'.'});
    if(insWorstFCR&&insWorstFCR!==insBestFCR)panels.push({type:'warn',title:'Weakest Period',body:'Lowest FCR: <strong>'+insWorstFCR.rate.toFixed(1)+'%</strong> — '+keyLabel(insWorstFCR.k,mode)+'.'});
    if(insAvgComp>=80)panels.push({type:'good',title:'Compliance On Track',body:'Avg compliance <strong>'+insAvgComp.toFixed(1)+'%</strong> — fields completed consistently.'});
    else panels.push({type:'bad',title:'Compliance Gap',body:'Avg compliance <strong>'+insAvgComp.toFixed(1)+'%</strong> — <strong>'+insCompGap.toFixed(1)+'%</strong> of tickets missing Action Taken / Reason / Sub Reason.'});
    if(insAvgCont>=85)panels.push({type:'good',title:'Strong Containment',body:'Containment averaging <strong>'+insAvgCont.toFixed(1)+'%</strong>.'});
    else panels.push({type:'warn',title:'Containment Below Target',body:'Containment at <strong>'+insAvgCont.toFixed(1)+'%</strong> — review escalation triggers.'});
    if(insEscSpike>0)panels.push({type:'warn',title:'Escalation Spikes',body:'<strong>'+insEscSpike+'</strong> period(s) had escalation >50% of Decagon volume.'});
    if(insRepPct>10)panels.push({type:'warn',title:'High Repeat AI Rate',body:insRepPct.toFixed(1)+'% of tickets had multiple AI interactions.'});
    var ipHtml='<div style="display:flex;flex-direction:column;gap:10px">';
    panels.forEach(function(p){var bg={good:'#f0fdf4',warn:'#fffbeb',bad:'#fef2f2',info:'#f0fafa'}[p.type]||'#f8fffe';var br={good:'#bbf7d0',warn:'#fde68a',bad:'#fecaca',info:'#cce8e8'}[p.type]||'#cce8e8';ipHtml+='<div style="background:'+bg+';border:1px solid '+br+';border-radius:10px;padding:12px 14px"><div style="display:flex;align-items:center;gap:6px;margin-bottom:4px">'+insIcon(p.type)+'<span style="font-size:12px;font-weight:700;color:#1e293b">'+p.title+'</span></div><div style="font-size:12px;color:#475569;line-height:1.5">'+p.body+'</div></div>';});
    ipHtml+='</div>';
    var insPanel='<div style="width:280px;flex-shrink:0"><div style="font-size:11px;font-weight:700;color:#0d9488;text-transform:uppercase;letter-spacing:0.08em;margin-bottom:10px">💡 Key Insights</div>'+ipHtml+'<div style="margin-top:10px;font-size:10px;color:#94a3b8;font-style:italic">Based on visible period data</div></div>';
    h+='<div style="display:flex;gap:20px;align-items:flex-start"><div style="flex:1;min-width:0">'+tbl+'</div>'+insPanel+'</div>';

        el.innerHTML=h;
    window._tMode=m=>{mode=m;offset=0;render();};
    window._tNav=dir=>{const ks=Object.keys(getBuckets()).sort();const ps2=mode==='daily'?7:mode==='weekly'?4:ks.length;const mo=Math.max(0,Math.floor((ks.length-ps2)/ps2));offset=Math.max(0,Math.min(mo,offset+dir));render();};
  }
  render();
}

function exportPDF() {
  if (!STATE.filteredTickets.size) { showToast('Upload data first','error'); return; }
  showToast('Generating PDF — please wait...', 'info');

  const { jsPDF } = window.jspdf;
  const doc = new jsPDF({ orientation: 'landscape', unit: 'mm', format: 'a4' });
  const pageW = 297, pageH = 210;
  const now = new Date().toLocaleDateString('en-GB');

  const tabs = ['kpis','effectiveness','compliance','defects','reasons','fcr','recontact','executive','tickets'];
  const tabLabels = { kpis:'Executive KPIs', effectiveness:'Decagon Effectiveness', compliance:'Decagon Compliance', defects:'System Defects', reasons:'Reason Analysis', fcr:'FCR Analysis', recontact:'Re-contact Analysis','decagon-report':'Decagon Report', executive:'Executive Summary', tickets:'Master Tickets' };

  const originalActive = document.querySelector('.nav-item.active')?.dataset?.tab || 'kpis';
  let pageNum = 0;

  async function captureTabs(tabList) {
    for (const tabId of tabList) {
      document.querySelectorAll('.nav-item').forEach(n => n.classList.remove('active'));
      document.querySelectorAll('.tab-panel').forEach(p => p.classList.remove('active'));
      const navItem = document.querySelector('.nav-item[data-tab="' + tabId + '"]');
      const panel = document.getElementById('tab-' + tabId);
      if (!navItem || !panel) continue;
      navItem.classList.add('active');
      panel.classList.add('active');
      document.getElementById('topbarTitle').textContent = tabLabels[tabId] || tabId;
      if (tabId === 'fcr' && !STATE.fcrBuilt) { buildFCRTab(); STATE.fcrBuilt = true; }
      await new Promise(r => setTimeout(r, 500));

      const el = document.getElementById('mainWrapper');
      const prevOverflow = el.style.overflow;
      el.style.overflow = 'visible';
      const canvas = await html2canvas(el, { scale: 1.2, useCORS: true, logging: false, backgroundColor: '#f1f5f9', windowWidth: 1400, scrollX: 0, scrollY: 0 });
      el.style.overflow = prevOverflow;
      const imgData = canvas.toDataURL('image/jpeg', 0.85);

      if (pageNum > 0) doc.addPage();
      doc.setFillColor(15, 23, 42);
      doc.rect(0, 0, pageW, 10, 'F');
      doc.setTextColor(255, 255, 255);
      doc.setFontSize(8);
      doc.setFont('helvetica', 'bold');
      doc.text('WAY-DECAGON DASHBOARD  |  ' + (tabLabels[tabId] || tabId).toUpperCase(), 6, 7);
      doc.text(now, pageW - 6, 7, { align: 'right' });

      const imgH = (canvas.height / canvas.width) * pageW;
      const finalH = Math.min(imgH, pageH - 12);
      doc.addImage(imgData, 'JPEG', 0, 11, pageW, finalH);
      pageNum++;
    }

    document.querySelectorAll('.nav-item').forEach(n => n.classList.remove('active'));
    document.querySelectorAll('.tab-panel').forEach(p => p.classList.remove('active'));
    const origNav = document.querySelector('.nav-item[data-tab="' + originalActive + '"]');
    if (origNav) origNav.classList.add('active');
    const origPanel = document.getElementById('tab-' + originalActive);
    if (origPanel) origPanel.classList.add('active');

    doc.save('way_decagon_' + now.replace(/\//g, '-') + '.pdf');
    showToast('PDF exported — ' + pageNum + ' pages', 'success');
  }

  captureTabs(tabs).catch(err => { showToast('PDF error: ' + err.message, 'error'); console.error(err); });
}
function exportSummary(){
  if(!STATE.filteredTickets.size)return;
  const m=computeMetrics(STATE.filteredTickets);
  const text=`WAY-DECAGON EXECUTIVE SUMMARY\nGenerated: ${new Date().toLocaleString()}\n${'='.repeat(50)}\nTotal Calls (Human + Decagon): ${fmt.num(trueTotalCalls)}\nCalls Routed to Decagon: ${fmt.num(m.decagonTickets)}\nCalls Handled by Decagon Alone: ${fmt.num(m.decagonOnlyCount)}\nEscalated to CS: ${fmt.num(m.csAssistedCount)}\nDecagon FCR: ${fmt.pct(m.fcrRate)}\nContainment Rate: ${fmt.pct(m.containmentRate)}\nCompliance Rate (Order/User level): ${fmt.pct(m.complianceRate)}\nCompliance Failures (Order/User level): ${fmt.num(m.complianceFailures)}\nUnique Orders/Users: ${fmt.num(m.complianceGroups)}\nTicket-Level Compliance Rate: ${fmt.pct(m.ticketComplianceRate)}\nStatus Not Closed: ${fmt.num(m.statusNotClosed)}\n`;
  const a=document.createElement('a');a.href=URL.createObjectURL(new Blob([text],{type:'text/plain'}));a.download='way_decagon_summary.txt';a.click();
  showToast('Summary exported','success');
}

// ── NAV ──
// ── FCR TAB ──
function buildFCRTab(){
  const dec=[...STATE.filteredTickets.values()].filter(t=>t.isDecagonTicket);
  const met=dec.filter(t=>t.fcrAchieved);
  const notMet=dec.filter(t=>!t.fcrAchieved);
  const trueFCR=dec.filter(t=>t.fcrAchieved&&t.compliant);
  const csAssisted=dec.filter(t=>t.csAssisted);
  const multiAI=dec.filter(t=>!t.csAssisted&&t.aiInteractionCount>1);
  const shortInt=dec.filter(t=>!t.csAssisted&&t.aiInteractionCount===1&&t.shortIntervalFlag);
  const sr=document.getElementById('fcrSummaryRow');
  if(!sr)return;
  const fcrRate=dec.length?(met.length/dec.length*100).toFixed(1):0;
  const trueFCRRate=dec.length?(trueFCR.length/dec.length*100).toFixed(1):0;
  sr.innerHTML=`
    <div class="kpi-card" style="border-top:3px solid #059669"><div class="kpi-label">FCR MET</div><div class="kpi-val" style="color:#059669">${fmt.num(met.length)}</div><div class="kpi-sub">${fcrRate}% of ${fmt.num(dec.length)} tickets</div></div>
    <div class="kpi-card" style="border-top:3px solid #dc2626"><div class="kpi-label">FCR NOT MET</div><div class="kpi-val" style="color:#dc2626">${fmt.num(notMet.length)}</div><div class="kpi-sub">${(100-fcrRate).toFixed(1)}% of total</div></div>
    <div class="kpi-card" style="border-top:3px solid #7c3aed"><div class="kpi-label">TRUE FCR (FCR + Compliant)</div><div class="kpi-val" style="color:#7c3aed">${fmt.num(trueFCR.length)}</div><div class="kpi-sub">${trueFCRRate}% of total</div></div>
    <div class="kpi-card" style="border-top:3px solid #f59e0b"><div class="kpi-label">CS ASSISTED</div><div class="kpi-val" style="color:#f59e0b">${fmt.num(csAssisted.length)}</div><div class="kpi-sub">${dec.length?(csAssisted.length/dec.length*100).toFixed(1):0}% of total</div></div>
    <div class="kpi-card" style="border-top:3px solid #3b82f6"><div class="kpi-label">MULTIPLE AI CALLS</div><div class="kpi-val" style="color:#3b82f6">${fmt.num(multiAI.length)}</div><div class="kpi-sub">${dec.length?(multiAI.length/dec.length*100).toFixed(1):0}% of total</div></div>`;
  const{text,grid}=getCC();
  const base={responsive:true,maintainAspectRatio:true,interaction:{mode:'index',intersect:false},plugins:{legend:{labels:{color:text,font:{size:11}}},tooltip:{backgroundColor:'#fff',titleColor:'#0f172a',bodyColor:'#475569',borderColor:'#e2e8f0',borderWidth:1}},scales:{x:{ticks:{color:text,font:{size:10},maxRotation:45},grid:{color:grid}},y:{ticks:{color:text,font:{size:10}},grid:{color:grid},beginAtZero:true}}};
  const dateMap={};
  dec.forEach(t=>{const d=t.dateBucket||'Unknown';if(!dateMap[d])dateMap[d]={met:0,total:0};dateMap[d].total++;if(t.fcrAchieved)dateMap[d].met++;});
  const dates=Object.keys(dateMap).sort();
  const fcrRates=dates.map(d=>dateMap[d].total?(dateMap[d].met/dateMap[d].total*100).toFixed(1):0);
  setTimeout(()=>{
    const ctx3=document.getElementById('fcrTrendChart');
    if(ctx3){if(ctx3._chart)ctx3._chart.destroy();ctx3._chart=new Chart(ctx3,{type:'line',data:{labels:dates,datasets:[{label:'FCR %',data:fcrRates,borderColor:'#059669',backgroundColor:'rgba(5,150,105,0.1)',fill:true,tension:0.4,pointRadius:3}]},options:{...base,scales:{x:base.scales.x,y:{...base.scales.y,ticks:{...base.scales.y.ticks,callback:v=>v+'%'},max:100}}}});}
    const ctx4=document.getElementById('fcrFailChart');
    if(ctx4){if(ctx4._chart)ctx4._chart.destroy();ctx4._chart=new Chart(ctx4,{type:'doughnut',data:{labels:['CS Assisted','Multiple AI Calls','Short Interval'],datasets:[{data:[csAssisted.length,multiAI.length,shortInt.length],backgroundColor:['#f59e0b','#3b82f6','#ef4444'],borderWidth:2}]},options:{...base,plugins:{...base.plugins,legend:{display:true,position:'bottom'}}}});}
    const excl2=new Set(CONFIG.EXCLUDED_REASONS);
    const rgMap={};
    dec.forEach(t=>{
      const r=(t.displayReason&&t.displayReason.trim())?t.displayReason.trim():'(No Reason)';
      const sr2=(t.subReason&&t.subReason.trim()&&!excl2.has(t.subReason.trim().toLowerCase()))?t.subReason.trim():'(No Sub Reason)';
      if(!rgMap[r])rgMap[r]={met:0,notMet:0,subs:{}};
      t.fcrAchieved?rgMap[r].met++:rgMap[r].notMet++;
      if(!rgMap[r].subs[sr2])rgMap[r].subs[sr2]={met:0,notMet:0};
      t.fcrAchieved?rgMap[r].subs[sr2].met++:rgMap[r].subs[sr2].notMet++;
    });
    const rtDiv=document.getElementById('fcrReasonTable');
    if(rtDiv){
      const html2=Object.entries(rgMap).sort((a,b)=>(b[1].met+b[1].notMet)-(a[1].met+a[1].notMet)).map(([reason,g])=>{
        const uid='u'+Math.random().toString(36).slice(2,8);
        const subs=Object.entries(g.subs).sort((a,b)=>(b[1].met+b[1].notMet)-(a[1].met+a[1].notMet)).map(([sr3,v])=>`<tr class="sr_${uid}" style="display:none;background:#f8fafc"><td style="padding-left:2rem;color:#64748b;font-size:11px">\u21b3 ${sr3}</td><td><span class="fcr-met-badge">${v.met}</span></td><td><span class="fcr-fail-badge">${v.notMet}</span></td><td>${v.met+v.notMet}</td></tr>`).join('');
        return `<tr style="cursor:pointer;font-weight:500" onclick="this.parentNode.querySelectorAll('.sr_${uid}').forEach(s=>s.style.display=s.style.display==='none'?'':'none');this.querySelector('span.arr').textContent=this.querySelector('span.arr').textContent==='\u25b6'?'\u25bc':'\u25b6'"><td><span class="arr" style="font-size:10px;color:#94a3b8;margin-right:4px">\u25b6</span>${reason} <span style="font-size:10px;color:#94a3b8;font-weight:400">(${Object.keys(g.subs).length})</span></td><td><span class="fcr-met-badge">${g.met}</span></td><td><span class="fcr-fail-badge">${g.notMet}</span></td><td>${g.met+g.notMet}</td></tr>${subs}`;
      }).join('');
      rtDiv.innerHTML=`<table class="fcr-reason-table" style="width:100%"><thead><tr><th>Reason (click to expand)</th><th>FCR Met</th><th>Not Met</th><th>Total</th></tr></thead><tbody>${html2}</tbody></table>`;
    }
    if(STATE.fcrDrillTable){try{STATE.fcrDrillTable.destroy();}catch(e){}STATE.fcrDrillTable=null;}
    const statusFilter=document.getElementById('fcrFilterStatus')?.value||'';
    const reasonFilter=document.getElementById('fcrFilterReason')?.value||'';
    let filtered=dec;
    if(statusFilter==='met')filtered=filtered.filter(t=>t.fcrAchieved);
    if(statusFilter==='notmet')filtered=filtered.filter(t=>!t.fcrAchieved);
    if(reasonFilter==='cs')filtered=filtered.filter(t=>t.csAssisted);
    if(reasonFilter==='multi')filtered=filtered.filter(t=>!t.csAssisted&&t.aiInteractionCount>1);
    if(reasonFilter==='short')filtered=filtered.filter(t=>!t.csAssisted&&t.aiInteractionCount===1&&t.shortIntervalFlag);
    const failReason=t=>{if(t.fcrAchieved)return'—';if(t.csAssisted)return'CS Assisted';if(t.aiInteractionCount>1)return'Multiple AI Calls';if(t.shortIntervalFlag)return'Short Interval';return'Other';};
    STATE.fcrDrillTable=new DataTable('#fcrDrillTable',{data:filtered,destroy:true,deferRender:true,pageLength:25,columns:[
      {title:'Ticket ID',data:'ticketId'},{title:'Date',data:'dateBucket'},
      {title:'FCR',data:'fcrAchieved',render:d=>d?'<span class="fcr-met-badge">MET</span>':'<span class="fcr-fail-badge">NOT MET</span>'},
      {title:'Failure Reason',data:null,render:(d,t,r)=>failReason(r)},
      {title:'Reason',data:'reason',render:d=>d||'\u2014'},{title:'Sub Reason',data:'subReason',render:d=>d||'\u2014'},
      {title:'Status',data:'status',render:d=>d||'\u2014'},{title:'AI Interactions',data:'aiInteractionCount'}
    ]});
    const fcrApplyBtn=document.getElementById('fcrApplyFilter');
    if(fcrApplyBtn)fcrApplyBtn.onclick=buildFCRTab;
  },50);
}

function setupNav(){
  const TITLES={upload:'Data Source',kpis:'Executive KPIs',trends:'Trend Analysis',effectiveness:'Decagon Effectiveness',fcr:'FCR Analysis',compliance:'Decagon Compliance',defects:'System Defects',reasons:'Reason Analysis',executive:'Executive Summary',validation:'Data Validation',tickets:'Master Tickets',recontact:'Re-contact Analysis','decagon-report':'Decagon Report'};
  document.querySelectorAll('.nav-item').forEach(item=>{
    item.addEventListener('click',()=>{
      document.querySelectorAll('.nav-item').forEach(n=>n.classList.remove('active'));
      document.querySelectorAll('.tab-panel').forEach(p=>p.classList.remove('active'));
      item.classList.add('active');
      const tab=item.dataset.tab;
      document.getElementById('tab-'+tab)?.classList.add('active');
      document.getElementById('topbarTitle').textContent=TITLES[tab]||tab;
      if(tab==='fcr'){if(!STATE.fcrBuilt){buildFCRTab();STATE.fcrBuilt=true;}}
      if(tab==='trends'){buildTrendsTab();}
    });
  });
  document.getElementById('sidebarToggle').addEventListener('click',()=>{
    document.getElementById('sidebar').classList.toggle('collapsed');
    document.getElementById('mainWrapper').classList.toggle('expanded');
  });
}


// ── DECAGON REPORT ──
const DR = { crmRows:null, crmFile:null, decRows:null, decFile:null, decSheets:0, result:null, tables:{} };
function drNormPhone(p){ if(p==null)return null; const s=String(p).replace(/[^\d]/g,''); if(!s)return null; if(s.length===10)return '+1'+s; if(s.length===11&&s[0]==='1')return '+1'+s.slice(1); return '+'+s; }
function drIsQA(n){ if(!n)return false; if(n.startsWith('+91'))return true; const d=n.replace(/[^\d]/g,''); return d==='14085983338'||d==='4085983338'; }
async function drReadDecagonWorkbook(file){
  return new Promise((resolve,reject)=>{ const r=new FileReader();
    r.onload=e=>{ try{ const wb=XLSX.read(new Uint8Array(e.target.result),{type:'array',cellDates:true}); const all=[];
      wb.SheetNames.forEach(sn=>{ const rows=XLSX.utils.sheet_to_json(wb.Sheets[sn],{defval:''});
        rows.forEach(row=>{ all.push({ _sheet:sn, phone:row['Phone Number']||row['Phone']||'', duration:row['Duration']||'',
          escDefl:String(row['Escalated/Deflected']||'').trim(), csat:row['CSAT']||'', timestamp:row['Timestamp (PDT)']||row['Timestamp']||'',
          medium:String(row['Medium']||'').trim(), aop:String(row['AOP']||'').trim(), link:row['Conversation Link']||'' }); }); });
      resolve({rows:all,sheetCount:wb.SheetNames.length}); }catch(err){reject(err);} };
    r.onerror=()=>reject(new Error('read failed')); r.readAsArrayBuffer(file); });
}
function setupDecagonReport(){
  const dzC=document.getElementById('drDropCRM'), dzD=document.getElementById('drDropDec');
  const fiC=document.getElementById('drFileCRM'), fiD=document.getElementById('drFileDec'); if(!dzC||!dzD)return;
  document.getElementById('drBrowseCRM')?.addEventListener('click',e=>{e.stopPropagation();fiC.click();});
  document.getElementById('drBrowseDec')?.addEventListener('click',e=>{e.stopPropagation();fiD.click();});
  const bind=(dz,fi,cb)=>{
    dz.addEventListener('click',e=>{const t=e.target; if(t===dz||t.classList.contains('drop-icon')||t.classList.contains('drop-title')||t.classList.contains('drop-sub')||t.classList.contains('drop-format'))fi.click();});
    dz.addEventListener('dragover',e=>{e.preventDefault();dz.classList.add('drag-over');});
    dz.addEventListener('dragleave',()=>dz.classList.remove('drag-over'));
    dz.addEventListener('drop',e=>{e.preventDefault();dz.classList.remove('drag-over'); if(e.dataTransfer.files[0])cb(e.dataTransfer.files[0]);});
    fi.addEventListener('change',e=>{if(e.target.files[0])cb(e.target.files[0]);});
  };
  bind(dzC,fiC,drHandleCRM); bind(dzD,fiD,drHandleDec);
  document.getElementById('drRunBtn').addEventListener('click',drRunReport);
}
function drHandleCRM(file){ showToast('Reading CRM…','info'); const done=rows=>{DR.crmRows=rows;DR.crmFile=file.name;drUpdateStatus();showToast('CRM · '+rows.length+' rows','success');};
  const n=file.name.toLowerCase();
  if(n.endsWith('.xlsx')||n.endsWith('.xls'))readXLSX(file).then(done).catch(err=>showToast('CRM error: '+err.message,'error'));
  else Papa.parse(file,{header:true,skipEmptyLines:true,complete:r=>done(r.data),error:err=>showToast('CRM error: '+err.message,'error')});
}
function drHandleDec(file){ showToast('Reading Decagon…','info');
  drReadDecagonWorkbook(file).then(res=>{DR.decRows=res.rows;DR.decFile=file.name;DR.decSheets=res.sheetCount;drUpdateStatus();showToast('Decagon · '+res.sheetCount+' sheets · '+res.rows.length+' rows','success');})
    .catch(err=>showToast('Decagon error: '+err.message,'error'));
}
function drUpdateStatus(){ document.getElementById('drStatus').style.display='block';
  document.getElementById('drCrmFile').textContent=DR.crmFile||'—';
  document.getElementById('drCrmRows').textContent=DR.crmRows?DR.crmRows.length.toLocaleString():0;
  document.getElementById('drDecFile').textContent=DR.decFile||'—';
  document.getElementById('drDecSheets').textContent=DR.decSheets||0;
  document.getElementById('drDecRows').textContent=DR.decRows?DR.decRows.length.toLocaleString():0;
}
function drRunReport(){
  if(!DR.decRows){showToast('Load Decagon file first','error');return;}
  const nonAddr=new Set((document.getElementById('drNonAddr').value||'').split(',').map(s=>s.trim().toLowerCase()).filter(Boolean));
  const voice=DR.decRows.filter(r=>r.medium.toLowerCase()==='voice');
  const qaRows=[], realRows=[];
  voice.forEach(r=>{r._normPhone=drNormPhone(r.phone); if(drIsQA(r._normPhone))qaRows.push(r); else realRows.push(r);});
  const isDefl=r=>r.escDefl.toLowerCase().includes('deflect'), isEsc=r=>r.escDefl.toLowerCase().includes('escalat');
  const deflected=realRows.filter(isDefl), escalated=realRows.filter(isEsc);
  const addressable=realRows.filter(r=>!nonAddr.has(r.aop.toLowerCase())); const addressableDefl=addressable.filter(isDefl);
  let matched=0, unmatched=0;
  if(DR.crmRows&&DR.crmRows.length){
    const sample=DR.crmRows[0]; const phoneCol=Object.keys(sample).find(k=>/phone/i.test(k))||null;
    const intCol=Object.keys(sample).find(k=>/^interaction$|interaction type/i.test(k))||null;
    if(phoneCol){ const crmSet=new Set();
      DR.crmRows.forEach(row=>{ if(intCol){const it=String(row[intCol]||'').trim(); if(it!=='AI-Agent Call'&&it!=='Call')return;}
        const np=drNormPhone(row[phoneCol]); if(np)crmSet.add(np); });
      realRows.forEach(r=>{r._crmMatch=!!(r._normPhone&&crmSet.has(r._normPhone)); if(r._crmMatch)matched++; else unmatched++;});
    }
  }
  DR.result={totalVoice:voice.length,qaExcluded:qaRows.length,real:realRows.length,deflected:deflected.length,escalated:escalated.length,
    addressable:addressable.length,addressableDefl:addressableDefl.length,matched,unmatched,qaRows,realRows,crmLoaded:!!DR.crmRows};
  drRender();
}
function drKpi(label,val,sub,icon,color){ return '<div class="kpi-card kpi-'+color+'"><div class="kpi-icon"><i class="fa-solid '+icon+'"></i></div><div class="kpi-body"><div class="kpi-label">'+label+'</div><div class="kpi-value">'+val+'</div>'+(sub?'<div class="kpi-sub">'+sub+'</div>':'')+'</div></div>'; }
function drRender(){
  const r=DR.result; document.getElementById('drResults').style.display='block';
  const oP=pct(r.deflected,r.real), aP=pct(r.addressableDefl,r.addressable), eP=pct(r.escalated,r.real);
  const kpis=[
    drKpi('Total Voice Conversations',fmt.num(r.totalVoice),'From all sheets','fa-phone','blue'),
    drKpi('QA Excluded',fmt.num(r.qaExcluded),'+91 & 4085983338','fa-shield-halved','muted'),
    drKpi('Real Customer Calls',fmt.num(r.real),'Voice · QA removed','fa-user','cyan'),
    drKpi('Deflected',fmt.num(r.deflected),fmt.pct(oP)+' of real','fa-circle-check','green'),
    drKpi('Escalated',fmt.num(r.escalated),fmt.pct(eP)+' of real','fa-arrow-up-right-from-square','orange'),
    drKpi('Overall Deflection Rate',fmt.pct(oP),r.deflected+' / '+r.real,'fa-percent','green'),
    drKpi('Addressable Deflection Rate',fmt.pct(aP),r.addressableDefl+' / '+r.addressable+' addressable','fa-bullseye','teal')
  ];
  if(r.crmLoaded){ kpis.push(drKpi('CRM Matched',fmt.num(r.matched),fmt.pct(pct(r.matched,r.real)),'fa-link','purple'));
    kpis.push(drKpi('CRM Gap',fmt.num(r.unmatched),fmt.pct(pct(r.unmatched,r.real)),'fa-link-slash','red')); }
  document.getElementById('drKpiGrid').innerHTML=kpis.join('');

  const dateAgg={};
  r.realRows.forEach(row=>{ const d=row._sheet; if(!dateAgg[d])dateAgg[d]={date:d,total:0,deflected:0,escalated:0}; dateAgg[d].total++;
    if(row.escDefl.toLowerCase().includes('deflect'))dateAgg[d].deflected++;
    if(row.escDefl.toLowerCase().includes('escalat'))dateAgg[d].escalated++; });
  const dateRows=Object.values(dateAgg).sort((a,b)=>a.date.localeCompare(b.date));

  ['_dr_date','_dr_qa','_dr_all'].forEach(k=>{ if(DR.tables[k]){try{DR.tables[k].destroy();}catch(e){} DR.tables[k]=null;} });
  $('#drDateTable').empty(); $('#drQATable').empty(); $('#drAllTable').empty();

  DR.tables._dr_date=$('#drDateTable').DataTable({
    data:dateRows.map(d=>[d.date,d.total,d.deflected,d.escalated,pct(d.deflected,d.total).toFixed(1)+'%']),
    columns:[{title:'Sheet Date'},{title:'Real Calls'},{title:'Deflected'},{title:'Escalated'},{title:'Deflection %'}],
    pageLength:25, order:[[0,'asc']], dom:'Bfrtip', buttons:['copy','csv'] });

  DR.tables._dr_qa=$('#drQATable').DataTable({
    data:r.qaRows.map(row=>[row._sheet,String(row.phone),row._normPhone||'—',row.escDefl||'—',row.aop||'—',String(row.timestamp||'—')]),
    columns:[{title:'Sheet'},{title:'Phone'},{title:'Normalized'},{title:'Esc/Defl'},{title:'AOP'},{title:'Timestamp'}],
    pageLength:10, dom:'Bfrtip', buttons:['copy','csv'] });

  DR.tables._dr_all=$('#drAllTable').DataTable({
    data:r.realRows.map(row=>[row._sheet,row._normPhone||String(row.phone),String(row.duration||''),row.escDefl||'—',row.aop||'—',String(row.csat||''),String(row.timestamp||''),
      r.crmLoaded?(row._crmMatch?'<span style="color:#059669">✓ Matched</span>':'<span style="color:#dc2626">✗ Gap</span>'):'—',
      row.link?'<a href="'+row.link+'" target="_blank">Open</a>':'—']),
    columns:[{title:'Date'},{title:'Phone'},{title:'Duration'},{title:'Esc/Defl'},{title:'AOP'},{title:'CSAT'},{title:'Timestamp'},{title:'CRM Match'},{title:'Link'}],
    pageLength:25, dom:'Bfrtip', buttons:['copy','csv'] });
}


// ── INIT ──
document.addEventListener('DOMContentLoaded',()=>{
  setupUpload();setupNav();setupDecagonReport();
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
