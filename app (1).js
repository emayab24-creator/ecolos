
// ====== State & mappings ======
window.APP = window.APP || {
  DealsA: { rows: [] },
  mappings: {
    departments: {'Эколос Алматы':[],'ЗИО Ecolos':[],'Ecolos Engineering':[]},
    groups: {
      OP: ['Аблай Каракожаев','Адиль Аманов','Асан Тортаев','Кристина Гайдар','Мадина Абатова','Арслан Мурат','Азамат Байкуатов','Акбота Кудайбергенова','Асылбек','Темірлан Мәкен','Наталья Клюшина','Александр Лукерин'],
      MPO:['Александр Венедиктов','Диана Джакупова','Дмитрий Коваль','Евгений Николаев','Елена Зинкина','Леонид Крупин','Максат Садвакасов','Нургуль Олжабаева','Нурсулу Бопишева','Татьяна Чибурун','Фариза Темирбекова'],
      Tender:['Арслан Мурат']
    },
    allowedByPerson: {},
    stageOrder: ['Новая','В работе','Выдано проектировщику','В экспертизе','Экспертиза пройдена','Идет тендер','Сбор данных/подготовка ТКП','ТКП отправлено','ТКП согласовано','Договор на согласовании','Договор подписан','Производство','Отгружено','ШМ и ПН'],
    stageGroups: {'Проектирование':['Новая','В работе','Выдано проектировщику','В экспертизе'],'Тендер':['Экспертиза пройдена','Идет тендер'],'Реализация':['Сбор данных/подготовка ТКП','ТКП отправлено','ТКП согласовано','Договор на согласовании','Договор подписан','Производство','Отгружено','ШМ и ПН']},
    deptByPerson: {}
  },
  lastAgg: null,
  lastDealsImport: null,
  lastTasksImport: null
};
(()=>{ // init dept & allowed
  const m=APP.mappings;
  const almaty=['Аблай Каракожаев','Адиль Аманов','Асан Тортаев','Кристина Гайдар','Мадина Абатова','Арслан Мурат'];
  const zio=['Азамат Байкуатов','Акбота Кудайбергенова','Асылбек','Темірлан Мәкен','Наталья Клюшина'];
  const eng=[...m.groups.MPO];
  m.departments['Эколос Алматы']=almaty; m.departments['ЗИО Ecolos']=zio; m.departments['Ecolos Engineering']=eng;
  m.groups.OP.forEach(p=> m.allowedByPerson[p]=m.stageOrder.slice(6));
  m.groups.MPO.forEach(p=> m.allowedByPerson[p]=m.stageOrder.slice(0,4));
  m.allowedByPerson['Арслан Мурат']=m.stageOrder.slice(4,6);
  [...almaty,...zio,...eng].forEach(p=> m.deptByPerson[p] = almaty.includes(p)?'Эколос Алматы': zio.includes(p)?'ЗИО Ecolos':'Ecolos Engineering');
})();

// ====== Helpers & parsers ======
const $ = s => document.querySelector(s);
const fmt = n => (n==null||isNaN(n))?'':new Intl.NumberFormat('ru-RU').format(n);
const toast = (msg,type='info') => { const el=$('#toast'); el.textContent=msg; el.style.display='block'; setTimeout(()=> el.style.display='none', 3000) };
const log = (obj,label='') => { const el=$('#logPanel'); const t=new Date().toLocaleString(); const pretty=typeof obj==='string'?obj:JSON.stringify(obj,null,2); el.innerHTML=`<b>${t}</b> ${label?('— '+label):''}<pre>${pretty}</pre>`+el.innerHTML };
function setPill(id, text, kind){ const el=$(id); el.textContent=text; el.className='pill ' + (kind||''); }
function showZeroStates(){ const empty=!APP.DealsA.rows||APP.DealsA.rows.length===0; $('#zeroDeals').style.display=empty?'grid':'none'; $('#zeroDeals2').style.display=empty?'grid':'none'; }

async function readFileSmart(file){ const ab=await file.arrayBuffer(); let txt=new TextDecoder('utf-8',{fatal:false}).decode(ab); if((txt.match(/\uFFFD/g)||[]).length>5){ try{ txt=new TextDecoder('windows-1251').decode(ab) }catch(e){} } if(txt.charCodeAt(0)===0xFEFF) txt=txt.slice(1); return txt }
function detectSep(header){ const seps=[';','\t',',','|']; let best=';',score=-1; for(const s of seps){ const cnt=(header.match(new RegExp('\\'+s,'g'))||[]).length; if(cnt>score){score=cnt; best=s} } return best }
function parseCSVText(text){ const lines=text.split(/\r?\n/).filter(l=>l.length>0); if(!lines.length) return []; const sep=detectSep(lines[0]);
  function splitQuoted(line){ const out=[]; let cur='',q=false; for(let i=0;i<line.length;i++){ const ch=line[i]; if(ch==='\"'){ if(q && line[i+1]==='\"'){cur+='\"'; i++;} else {q=!q} continue } if(ch===sep && !q){ out.push(cur); cur=''; continue } cur+=ch } out.push(cur); return out }
  const headers=splitQuoted(lines[0]).map(s=>s.trim().replace(/^\"|\"$/g,'')); const rows=[]; for(let i=1;i<lines.length;i++){ const parts=splitQuoted(lines[i]).map(s=>s.trim().replace(/^\"|\"$/g,'')); const r={}; headers.forEach((h,idx)=> r[h]=(parts[idx]??'').trim() ); if(Object.values(r).every(v=>!String(v).trim())) continue; rows.push(r) } return rows }

// Stages & names normalization
const canonStages={'новая':'Новая','new':'Новая','вработе':'В работе','в работе':'В работе','inprogress':'В работе','выданопроектировщику':'Выдано проектировщику','project':'Выдано проектировщику','вэкспертизе':'В экспертизе','экспертиза':'В экспертизе','экспертизапройдена':'Экспертиза пройдена','passedexpertise':'Экспертиза пройдена','идеттендер':'Идет тендер','тендер':'Идет тендер','сборданныхподготовкаткп':'Сбор данных/подготовка ТКП','сборданных/подготовкаткп':'Сбор данных/подготовка ТКП','предв':'Сбор данных/подготовка ТКП','ткпотправлено':'ТКП отправлено','отправленоткп':'ТКП отправлено','ткпсогласовано':'ТКП согласовано','согласованоткп':'ТКП согласовано','договорнасогласовании':'Договор на согласовании','согласованиедоговора':'Договор на согласовании','договорподписан':'Договор подписан','подписандоговор':'Договор подписан','производство':'Производство','отгружено':'Отгружено','shipment':'Отгружено','шмипн':'ШМ и ПН','шм и пн':'ШМ и ПН','шмишмпн':'ШМ и ПН','шмпн':'ШМ и ПН'};
function normKey(s){return String(s||'').toLowerCase().replace(/[^\p{L}\p{N}]+/gu,'')}
function normalizeStage(raw){ const k=normKey(raw); return canonStages[k]||raw }
const knownPeople=['Аблай Каракожаев','Адиль Аманов','Асан Тортаев','Кристина Гайдар','Мадина Абатова','Арслан Мурат','Азамат Байкуатов','Акбота Кудайбергенова','Асылбек','Темірлан Мәкен','Наталья Клюшина','Александр Лукерин','Александр Венедиктов','Диана Джакупова','Дмитрий Коваль','Евгений Николаев','Елена Зинкина','Леонид Крупин','Максат Садвакасов','Нургуль Олжабаева','Нурсулу Бопишева','Татьяна Чибурун','Фариза Темирбекова'];
function canonName(raw){ if(!raw) return ''; const s=String(raw).replace(/\s+/g,' ').trim(); if(knownPeople.includes(s)) return s; const parts=s.toLowerCase().split(' ').filter(Boolean); for(const p of knownPeople){ const kp=p.toLowerCase().split(' '); if(parts.length===2 && kp.length===2 && ((parts[0]===kp[0]&&parts[1]===kp[1])||(parts[0]===kp[1]&&parts[1]===kp[0]))){ return p } } return s }

// Deals normalize & merge
function normalizeDeals(rows){ if(!rows||!rows.length)return{rows:[],info:{mapped:{},ignored:0}}; const aliases={id:['id','id сделки','deal id','номер','ид','код','deal'],resp:['ответственный','менеджер','сотрудник','ответственный менеджер','owner','responsible','manager','assignee','мпо ответственный'],stage:['стадия','стадия сделки','этап','status','stage','pipeline'],created:['дата создания','создана','created','date created','создание','create time','created at'],updated:['дата изменения','изменена','updated','date updated','обновление','update time','updated at','modified']}; const norm=s=>String(s||'').toLowerCase().replace(/\s+/g,' ').trim(); const cols={}; Object.keys(rows[0]||{}).forEach(c=>cols[c]=norm(c)); const pick=keys=>{for(const[orig,n]of Object.entries(cols)){if(keys.some(a=>n.includes(a)))return orig}return null}; const cId=pick(aliases.id),cR=pick(aliases.resp),cS=pick(aliases.stage),cC=pick(aliases.created),cU=pick(aliases.updated); let ignored=0;
  const out=rows.map(r=>{ const stage=normalizeStage(r[cS]??r['Стадия сделки']??''); const resp=canonName((r[cR]??r['Ответственный']??r['МПО Ответственный']??'').trim()); if(!resp||!stage){ignored++}
    const iso=d=>{ if(!d)return null; const s=String(d); const m=s.match(/^(\d{2})\.(\d{2})\.(\d{4})(?:\s+(\d{2}):(\d{2})(?::(\d{2}))?)?/); if(m){return `${m[3]}-${m[2]}-${m[1]}${m[4]?(' '+m[4]+':'+m[5]+(m[6]?':'+m[6]:'') ):''}`} const ts=Date.parse(s); return isNaN(ts)?s:new Date(ts).toISOString().replace('T',' ').slice(0,19)};
    return {'ID сделки':(r[cId]||r['ID']||r['ID сделки']||'')||null,'Ответственный':resp,'Стадия сделки':stage,'Дата создания':iso(r[cC]||r['Дата создания']||null),'Дата изменения':iso(r[cU]||r['Дата изменения']||null),'Отдел':(APP.mappings.deptByPerson||{})[resp]||'—'} });
  return {rows:out, info:{mapped:{'ID сделки':cId||'—','Ответственный':cR||'—','Стадия сделки':cS||'—','Дата создания':cC||'—','Дата изменения':cU||'—'},ignored}}
}
function mergeDeals(existing,incoming){ const map=new Map(); (existing||[]).forEach(r=>{const id=r['ID сделки']; if(id)map.set(String(id),r)}); incoming.forEach(r=>{const id=r['ID сделки']; if(id){map.set(String(id),r)} else {(existing||[]).push(r)} }); const withId=Array.from(map.values()); const noId=(existing||[]).filter(r=>!r['ID сделки']); return withId.concat(noId) }

// ====== Export helpers ======
function exportTableXLS(id, filename){ const el=$(id); if(!el) return; const html='<!DOCTYPE html><html><head><meta charset="utf-8"></head><body>'+el.outerHTML+'</body></html>'; const blob=new Blob([html],{type:'application/vnd.ms-excel'}); const a=document.createElement('a'); a.href=URL.createObjectURL(blob); a.download=filename; a.click(); setTimeout(()=>URL.revokeObjectURL(a.href),1500) }
function exportMultipleTablesXLS(tables, filename){ const body = tables.map(sel=> $(sel)?.outerHTML || '').join('<br><br>'); const html='<!DOCTYPE html><html><head><meta charset="utf-8"></head><body>'+body+'</body></html>'; const blob=new Blob([html],{type:'application/vnd.ms-excel'}); const a=document.createElement('a'); a.href=URL.createObjectURL(blob); a.download=filename; a.click(); setTimeout(()=>URL.revokeObjectURL(a.href),1500) }

// ====== Dashboard builders ======
function buildKPI(rows){ const total=rows.length; const people=new Set(rows.map(r=>r['Ответственный']).filter(Boolean)); let mism=0; rows.forEach(r=>{ const p=r['Ответственный']; const s=r['Стадия сделки']; const al=APP.mappings.allowedByPerson[p]||[]; if(s&&al.length&&!al.includes(s)) mism++ }); $('#kpi').innerHTML=[['Сделок',total],['Сотрудников',people.size],['Несоответствий',mism]].map(([t,v])=>`<div class="card"><b>${t}</b><div style="font-size:18px;font-weight:800">${fmt(v)}</div></div>`).join('') }
function drawHBar(canvas, entries){ const ctx=canvas.getContext('2d'); const padL=120,padR=24,padT=16,padB=14; const H=canvas.height,W=canvas.width; ctx.clearRect(0,0,W,H); if(entries.length===0){ ctx.fillStyle='#e7ebf3'; ctx.fillText('Нет данных', 16, 24); return } const max=Math.max(1, ...entries.map(e=>e[1])); const barH=Math.max(12, Math.floor((H-padT-padB)/entries.length)-4); ctx.font='12px system-ui'; ctx.fillStyle='#e7ebf3'; entries.forEach((e,i)=>{ const [label,val]=e; const y=padT+i*(barH+4); const w=Math.round((W-padL-padR)*(val/max)); ctx.textAlign='right'; ctx.fillText(label,padL-6,y+barH-3); ctx.fillRect(padL,y,w,barH); ctx.textAlign='left'; ctx.fillText(String(val), padL+w+6, y+barH-3) }) }
function drawLine(canvas, items){ const ctx=canvas.getContext('2d'); const padL=36,padR=16,padT=12,padB=20; const H=canvas.height,W=canvas.width; ctx.clearRect(0,0,W,H); if(items.length===0){ ctx.fillStyle='#e7ebf3'; ctx.fillText('Нет данных',16,24); return } const max=Math.max(...items.map(i=>i[1])); const stepX=(W-padL-padR)/Math.max(1,items.length-1); ctx.strokeStyle='#445'; ctx.beginPath(); ctx.moveTo(padL,padT); ctx.lineTo(padL,H-padB); ctx.lineTo(W-padR,H-padB); ctx.stroke(); ctx.strokeStyle='#22d3ee'; ctx.beginPath(); items.forEach((it,idx)=>{ const x=padL+idx*stepX; const y=(H-padB)-(H-padT-padB)*(it[1]/max); if(idx===0)ctx.moveTo(x,y); else ctx.lineTo(x,y); ctx.fillStyle='#22d3ee'; ctx.beginPath(); ctx.arc(x,y,2,0,Math.PI*2); ctx.fill(); ctx.fillStyle='#e7ebf3'; ctx.textAlign='center'; ctx.fillText(String(it[1]), x, y-6) }); ctx.fillStyle='#e7ebf3'; ctx.textAlign='center'; items.forEach((it,idx)=>{ const x=padL+idx*stepX; ctx.fillText(it[0], x, H-4) }) }
function buildCharts(){ const rows=APP.DealsA.rows||[]; buildKPI(rows);
  const cm=$('#cMonth canvas')||(()=>{const c=document.createElement('canvas');$('#cMonth').appendChild(c);return c})(); const cd=$('#cDept canvas')||(()=>{const c=document.createElement('canvas');$('#cDept').appendChild(c);return c})(); const cs=$('#cStage canvas')||(()=>{const c=document.createElement('canvas');$('#cStage').appendChild(c);return c})(); const cp=$('#cPerson canvas')||(()=>{const c=document.createElement('canvas');$('#cPerson').appendChild(c);return c})();
  const ym=d=>d?`${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,'0')}`:null; const parseDate=v=>{if(!v)return null; const m=String(v).match(/^(\d{4})-(\d{2})-(\d{2})/); if(m)return new Date(+m[1],+m[2]-1,+m[3]); const m2=String(v).match(/^(\d{2})\.(\d{2})\.(\d{4})/); if(m2)return new Date(+m2[3],+m2[2]-1,+m2[1]); const ts=Date.parse(v); return isNaN(ts)?null:new Date(ts)};
  const months={}; rows.forEach(r=>{ const k=ym(parseDate(r['Дата создания']||r['Дата изменения'])); if(!k)return; months[k]=(months[k]||0)+1 });
  const mItems=Object.entries(months).filter(p=>p[1]>0).sort((a,b)=>a[0].localeCompare(b[0])); cm.width=$('#cMonth').clientWidth-2; cm.height=220; drawLine(cm,mItems);
  const byD=rows.reduce((acc,r)=>{ const k=r['Отдел']||'—'; acc[k]=(acc[k]||0)+1; return acc },{}); const dItems=Object.entries(byD).filter(e=>e[1]>0).sort((a,b)=>b[1]-a[1]); $('#cDept').style.height=(Math.max(3,dItems.length)*22+60)+'px'; cd.width=$('#cDept').clientWidth-2; cd.height=parseInt($('#cDept').style.height); drawHBar(cd,dItems);
  const st=Object.fromEntries(APP.mappings.stageOrder.map(s=>[s,0])); rows.forEach(r=>{ const s=r['Стадия сделки']; if(st[s]!=null) st[s]+=1 }); const sEntries=Object.entries(st).filter(e=>e[1]>0); $('#cStage').style.height=Math.min(280,(Math.max(6,sEntries.length)*22+60))+'px'; cs.width=$('#cStage').clientWidth-2; cs.height=parseInt($('#cStage').style.height); drawHBar(cs,sEntries);
  const byP=rows.reduce((acc,r)=>{ const k=r['Ответственный']||'—'; acc[k]=(acc[k]||0)+1; return acc },{}); const pEntries=Object.entries(byP).filter(e=>e[1]>0).sort((a,b)=>b[1]-a[1]).slice(0,20); $('#cPerson').style.height=Math.min(280,(Math.max(6,pEntries.length)*22+60))+'px'; cp.width=$('#cPerson').clientWidth-2; cp.height=parseInt($('#cPerson').style.height); drawHBar(cp,pEntries);
  APP.lastAgg = { months: mItems, depts: dItems, stages: sEntries, persons: pEntries };
  showZeroStates();
}
function exportDashboard(){ if(!APP.lastAgg){toast('Нет данных');return} const tbl=(title,rows)=>`<h3>${title}</h3><table><thead><tr><th>Ключ</th><th>Кол-во</th></tr></thead><tbody>${rows.map(r=>`<tr><td>${r[0]}</td><td>${r[1]}</td></tr>`).join('')}</tbody></table>`; const html='<!DOCTYPE html><html><head><meta charset="utf-8"></head><body>'+tbl('Месяцы',APP.lastAgg.months)+tbl('Отделы',APP.lastAgg.depts)+tbl('Стадии',APP.lastAgg.stages)+tbl('Сотрудники',APP.lastAgg.persons)+'</body></html>'; const blob=new Blob([html],{type:'application/vnd.ms-excel'}); const a=document.createElement('a'); a.href=URL.createObjectURL(blob); a.download='dashboard_data.xls'; a.click(); setTimeout(()=>URL.revokeObjectURL(a.href),1500) }

// ====== Week helpers ======
function isoWeek(d){ const date=new Date(Date.UTC(d.getFullYear(),d.getMonth(),d.getDate())); const day=(date.getUTCDay()+6)%7; date.setUTCDate(date.getUTCDate()-day+3); const firstThursday=new Date(Date.UTC(date.getUTCFullYear(),0,4)); const diff=date-firstThursday; return 1+Math.round(diff/604800000) }
function weekOfMonth(d){ const first=new Date(Date.UTC(d.getFullYear(), d.getMonth(), 1)); let dow=first.getUTCDay(); if(dow===0) dow=7; const offset=dow-1; return Math.floor((offset + d.getUTCDate() - 1)/7)+1 }
function monthName(m){ return ['Янв','Фев','Мар','Апр','Май','Июн','Июл','Авг','Сен','Окт','Ноя','Дек'][m-1] }
function labelFromMeta(meta){ const a=[]; if(meta.year&&meta.woy) a.push(`${meta.year}-W${String(meta.woy).padStart(2,'0')}`); if(meta.month&&meta.wom) a.push(`${monthName(meta.month)} • нед. ${meta.wom}`); return a.join(' | ')||'неделя не задана' }
function calcAutoMeta(rows){ // max(Дата изменения, Дата создания)
  const toDate = v => { if(!v) return null; const s=String(v); const m=s.match(/^(\d{4})-(\d{2})-(\d{2})/); if(m) return new Date(+m[1],+m[2]-1,+m[3]); const m2=s.match(/^(\d{2})\.(\d{2})\.(\d{4})/); if(m2) return new Date(+m2[3],+m2[2]-1,+m2[1]); const ts=Date.parse(s); return isNaN(ts)?null:new Date(ts) };
  let best=null; for(const r of (rows||[])){ const a=toDate(r['Дата изменения']); const b=toDate(r['Дата создания']); const d=a||b; if(d && (!best || d>best)) best=d; }
  if(!best) best=new Date();
  return { year: best.getFullYear(), month: best.getMonth()+1, wom: weekOfMonth(best), woy: isoWeek(best) };
}

// ====== Local storage for DEAL files ======
const DEAL_STORE_KEY='crm_deal_files_v1';
function loadDealFiles(){ try{ return JSON.parse(localStorage.getItem(DEAL_STORE_KEY)||'[]') }catch(e){ return [] } }
function saveDealFiles(arr){ localStorage.setItem(DEAL_STORE_KEY, JSON.stringify(arr)) }
function addDealFile(name, rows){ const files=loadDealFiles(); const meta=calcAutoMeta(rows); const id = Date.now().toString(36)+'_'+Math.random().toString(36).slice(2,6); files.push({id,name,rows,meta,uploaded:new Date().toISOString()}); saveDealFiles(files); refreshFileLists(); }
function refreshFileLists(){ const files=loadDealFiles(); const opt = f=> `<option value="${f.id}">${labelFromMeta(f.meta)} — ${f.name} — ${f.rows.length}</option>`;
  $('#fileList').innerHTML = files.map(opt).join('');
  $('#fileA').innerHTML = '<option value="">A: файл</option>'+files.map(opt).join('');
  $('#fileB').innerHTML = '<option value="">B: файл</option>'+files.map(opt).join('');
}
function getSelectedFiles(){ const ids=[...($('#fileList').selectedOptions||[])].map(o=>o.value); const files=loadDealFiles(); return files.filter(f=> ids.includes(f.id)) }
function fileById(id){ return loadDealFiles().find(f=>f.id===id) }
function updateSelectedMeta(auto=false){ const sel=getSelectedFiles(); if(sel.length!==1){ toast('Выбери ровно 1 файл в списке'); return } const f=sel[0]; if(auto){ f.meta = calcAutoMeta(f.rows) } else { f.meta = { year: parseInt($('#metaYear').value||'')||null, month: parseInt($('#metaMonth').value||'')||null, wom: parseInt($('#metaWOM').value||'')||null, woy: parseInt($('#metaWOY').value||'')||null } } const all=loadDealFiles().map(x=> x.id===f.id?f:x); saveDealFiles(all); refreshFileLists(); toast('Метки сохранены'); }
function fillMetaFromSelected(){ const sel=getSelectedFiles(); if(sel.length!==1){ $('#metaYear').value=''; $('#metaMonth').value=1; $('#metaWOM').value=1; $('#metaWOY').value=''; return } const m=sel[0].meta||{}; $('#metaYear').value=m.year||''; $('#metaMonth').value=m.month||1; $('#metaWOM').value=m.wom||1; $('#metaWOY').value=m.woy||''; }
function renameSelected(){ const sel=getSelectedFiles(); if(sel.length!==1){ toast('Выбери 1 файл для переименования'); return } const f=sel[0]; const name=prompt('Новое имя', f.name); if(!name) return; const all=loadDealFiles().map(x=> x.id===f.id?{...x,name}:x); saveDealFiles(all); refreshFileLists(); }
function deleteSelected(){ const ids=[...($('#fileList').selectedOptions||[])].map(o=>o.value); if(ids.length===0){ toast('Выбери файлы для удаления'); return } const left=loadDealFiles().filter(f=> !ids.includes(f.id)); saveDealFiles(left); refreshFileLists(); fillMetaFromSelected(); }

// ====== Comparison using saved files ======
function countsBy(rows, keyFn){ return rows.reduce((acc,r)=>{ const k=keyFn(r); if(!k) return acc; acc[k]=(acc[k]||0)+1; return acc },{}) }
function tableMatrix(rowKeys, colFiles, valueFn){
  const head = `<tr><th>Показатель</th>${colFiles.map(s=>`<th>${labelFromMeta(s.meta||{})}</th>`).join('')}</tr>`;
  const body = rowKeys.map(k=>{
    const cells = colFiles.map(s=>{ const v=valueFn(k, s); return `<td>${v?fmt(v):''}</td>` }).join('');
    return `<tr><th>${k}</th>${cells}</tr>`;
  }).join('');
  return `<table><thead>${head}</thead><tbody>${body}</tbody></table>`;
}
function compareMulti(){
  const sel=getSelectedFiles(); if(sel.length<2){ toast('Выбери 2+ файлов в списке'); return }
  const stageKeys = APP.mappings.stageOrder.slice();
  const deptKeys = Object.keys(APP.mappings.departments);
  const tblStages = tableMatrix(stageKeys, sel, (stage, f)=> countsBy(f.rows, r=> r['Стадия сделки'])[stage]||0);
  $('#multiStages').innerHTML = tblStages;
  const tblDepts = tableMatrix(deptKeys, sel, (dep, f)=> countsBy(f.rows, r=> r['Отдел'])[dep]||0);
  $('#multiDepts').innerHTML = tblDepts;
  const head = `<tr><th>Файл</th><th>Всего сделок</th></tr>`;
  const body = sel.map(s=> `<tr><td>${labelFromMeta(s.meta||{})} — ${s.name}</td><td>${fmt((s.rows||[]).length)}</td></tr>`).join('');
  $('#multiTotals').innerHTML = `<table id="multiTotalsTbl"><thead>${head}</thead><tbody>${body}</tbody></table>`;
  $('#exportMulti').onclick = ()=> exportMultipleTablesXLS(['#multiStages table','#multiDepts table','#multiTotalsTbl'],'compare_files_multi.xls');
}
function diffKV(a,b){ const keys=[...new Set([...Object.keys(a),...Object.keys(b)])]; const rows=keys.map(k=>{ const av=a[k]||0, bv=b[k]||0, d=bv-av; return [k,av,bv,d] }); return rows }
function fmtDelta(d){ if(d>0) return `<span class="delta-plus">+${d}</span>`; if(d<0) return `<span class="delta-minus">${d}</span>`; return '' }
function renderDiffTable(title, rows){ return `<table><thead><tr><th>${title}</th><th>A</th><th>B</th><th>Δ</th></tr></thead><tbody>${rows.map(r=>`<tr><td>${r[0]}</td><td>${r[1]}</td><td>${r[2]}</td><td>${fmtDelta(r[3])}</td></tr>`).join('')}</tbody></table>` }
function compareRows(A,B){
  const byIdA=new Map(), byIdB=new Map(); A.forEach(r=> byIdA.set(String(r['ID сделки']||''), r)); B.forEach(r=> byIdB.set(String(r['ID сделки']||''), r));
  const commonIds=[...byIdA.keys()].filter(id=> id && byIdB.has(id));
  const moves = commonIds.map(id=>{ const a=byIdA.get(id), b=byIdB.get(id); const from=a['Стадия сделки']||'', to=b['Стадия сделки']||''; const resp=b['Ответственный']||a['Ответственный']||''; const dept=b['Отдел']||a['Отдел']||''; return {id, from, to, resp, dept} }).filter(x=> x.from!==x.to);
  return {
    stages: diffKV(countsBy(A,r=>r['Стадия сделки']), countsBy(B,r=>r['Стадия сделки'])),
    depts: diffKV(countsBy(A,r=>r['Отдел']), countsBy(B,r=>r['Отдел'])),
    persons: diffKV(countsBy(A,r=>r['Ответственный']), countsBy(B,r=>r['Ответственный'])).sort((x,y)=> Math.abs(y[3])-Math.abs(x[3])).slice(0,30),
    moves
  }
}
function renderMoves(moves){ if(!moves.length) return '<div class="small">Перемещений не найдено</div>'; return `<table id="cmpMovesTbl"><thead><tr><th>ID</th><th>Было</th><th>Стало</th><th>Ответственный (тек.)</th><th>Отдел (тек.)</th></tr></thead><tbody>${moves.map(m=>`<tr><td>${m.id}</td><td>${m.from}</td><td>${m.to}</td><td>${m.resp}</td><td>${m.dept}</td></tr>`).join('')}</tbody></table>` }
function compareAB(){ const Aid=$('#fileA').value, Bid=$('#fileB').value; if(!Aid||!Bid){ toast('Выбери A и B'); return } const A=fileById(Aid); const B=fileById(Bid); const res=compareRows(A.rows||[],B.rows||[]); $('#cmpStages').innerHTML=renderDiffTable('Стадия',res.stages); $('#cmpDept').innerHTML=renderDiffTable('Отдел',res.depts); $('#cmpPersons').innerHTML=renderDiffTable('Сотрудник',res.persons); $('#cmpMoves').innerHTML=renderMoves(res.moves); $('#exportCompare').onclick=()=> exportMultipleTablesXLS(['#cmpStages table','#cmpDept table','#cmpPersons table','#cmpMovesTbl'],'compare_AB_files.xls') }

// ====== Mismatch & Stale ======
function stageClass(s){ const g=APP.mappings.stageGroups; if(g['Проектирование'].includes(s)) return 'col-proj'; if(g['Тендер'].includes(s)) return 'col-tender'; return 'col-real'; }
function personNameClass(p){ const G=APP.mappings.groups; if(G.Tender.includes(p)) return 'name-tender'; if(G.MPO.includes(p)) return 'name-mpo'; return 'name-op'; }

function buildMismatch(){ const base=APP.DealsA.rows||[]; const need=['Ответственный','Стадия сделки']; const missing=need.filter(c=>!base.length||base.every(r=>!r[c])); $('#mmWarn').textContent=missing.length?('Проверь колонки: '+missing.join(', ')):'';
  const depts=Object.keys(APP.mappings.departments);
  const headerTop=`<tr><th rowspan="2">Ответственный</th><th class="col-proj" colspan="${APP.mappings.stageGroups['Проектирование'].length}">Проектирование</th><th class="col-tender" colspan="${APP.mappings.stageGroups['Тендер'].length}">Тендер</th><th class="col-real" colspan="${APP.mappings.stageGroups['Реализация'].length}">Реализация</th><th rowspan="2">Итого</th></tr>`;
  const headerStages=`<tr>${APP.mappings.stageOrder.map((s,i)=>`<th class="${stageClass(s)}">${i+1}. ${s}</th>`).join('')}</tr>`;
  let body=''; const grand=Object.fromEntries(APP.mappings.stageOrder.map(s=>[s,0])); let grandTotal=0;
  depts.forEach(dep=>{ const ppl=APP.mappings.departments[dep]; let depTotals=Object.fromEntries(APP.mappings.stageOrder.map(s=>[s,0])); let depSum=0;
    ppl.forEach(p=>{ const rows=base.filter(r=>r['Ответственный']===p); const counts=Object.fromEntries(APP.mappings.stageOrder.map(s=>[s,0])); rows.forEach(r=>{ const st=r['Стадия сделки']; if(counts[st]!=null) counts[st]+=1 }); const total=Object.values(counts).reduce((a,b)=>a+b,0);
      depSum+=total; APP.mappings.stageOrder.forEach(s=>depTotals[s]+=counts[s]); const allowed=APP.mappings.allowedByPerson[p]||[];
      const tds=APP.mappings.stageOrder.map(s=>{ const v=counts[s]||0; const bad=v>0&&!allowed.includes(s); const cls=bad?'bad':(v>0?'gt0':''); return `<td class="${stageClass(s)}">${v>0?`<span class="${cls}">${fmt(v)}</span>`:''}</td>` }).join('');
      body+=`<tr><th class="${personNameClass(p)}">${p}</th>${tds}<th>${fmt(total)}</th></tr>`;
    });
    APP.mappings.stageOrder.forEach(s=>grand[s]+=depTotals[s]); grandTotal+=depSum;
    const depCells=APP.mappings.stageOrder.map(s=>`<td class="${stageClass(s)}"><b>${fmt(depTotals[s])}</b></td>`).join('');
    body+=`<tr><th>Итого ${dep}</th>${depCells}<th><b>${fmt(depSum)}</b></th></tr>`;
  });
  const grandCells=APP.mappings.stageOrder.map(s=>`<td class="${stageClass(s)}"><b>${fmt(grand[s])}</b></td>`).join('');
  const tfoot=`<tfoot><tr><th>Общий итог</th>${grandCells}<th><b>${fmt(grandTotal)}</b></th></tr></tfoot>`;
  $('#mmTable').innerHTML = `<div class="table-wrap"><table id="mismatchAll"><thead>${headerTop}${headerStages}</thead><tbody>${body}</tbody>${tfoot}</table></div>`;
  $('#exportMismatch').onclick = ()=> exportTableXLS('#mismatchAll','mismatch.xls');
  showZeroStates();
}

function buildStale(){ const daysInput=$('#staleDays');
  function apply(){ const N=parseInt(daysInput.value||'200',10); const now=new Date(); const base=(APP.DealsA.rows||[]).filter(r=>{ const d=new Date(r['Дата изменения']||r['Дата создания']); if(!d||isNaN(+d))return false; return (now - d)/(1000*60*60*24) > N });
    const depts=Object.keys(APP.mappings.departments);
    const headerTop=`<tr><th rowspan="2">Ответственный</th><th class="col-proj" colspan="${APP.mappings.stageGroups['Проектирование'].length}">Проектирование</th><th class="col-tender" colspan="${APP.mappings.stageGroups['Тендер'].length}">Тендер</th><th class="col-real" colspan="${APP.mappings.stageGroups['Реализация'].length}">Реализация</th><th rowspan="2">Итого</th></tr>`;
    const headerStages=`<tr>${APP.mappings.stageOrder.map((s,i)=>`<th class="${stageClass(s)}">${i+1}. ${s}</th>`).join('')}</tr>`;
    let body=''; const grand=Object.fromEntries(APP.mappings.stageOrder.map(s=>[s,0])); let grandTotal=0;
    depts.forEach(dep=>{ const ppl=APP.mappings.departments[dep]; let depTotals=Object.fromEntries(APP.mappings.stageOrder.map(s=>[s,0])); let depSum=0;
      ppl.forEach(p=>{ const rows=base.filter(r=>r['Ответственный']===p); const counts=Object.fromEntries(APP.mappings.stageOrder.map(s=>[s,0])); rows.forEach(r=>{ const st=r['Стадия сделки']; if(counts[st]!=null) counts[st]+=1 }); const total=Object.values(counts).reduce((a,b)=>a+b,0); depSum+=total; APP.mappings.stageOrder.forEach(s=>depTotals[s]+=counts[s]);
        const tds=APP.mappings.stageOrder.map(s=>{ const v=counts[s]||0; return `<td class="${stageClass(s)}">${v>0?`<span class="gt0">${fmt(v)}</span>`:''}</td>` }).join('');
        body+=`<tr><th class="${personNameClass(p)}">${p}</th>${tds}<th>${fmt(total)}</th></tr>`;
      });
      APP.mappings.stageOrder.forEach(s=>grand[s]+=depTotals[s]); grandTotal+=depSum;
      const depCells=APP.mappings.stageOrder.map(s=>`<td class="${stageClass(s)}"><b>${fmt(depTotals[s])}</b></td>`).join('');
      body+=`<tr><th>Итого ${dep}</th>${depCells}<th><b>${fmt(depSum)}</b></th></tr>`;
    });
    const grandCells=APP.mappings.stageOrder.map(s=>`<td class="${stageClass(s)}"><b>${fmt(grand[s])}</b></td>`).join('');
    const tfoot=`<tfoot><tr><th>Общий итог</th>${grandCells}<th><b>${fmt(grandTotal)}</b></th></tr></tfoot>`;
    $('#staleWrap').innerHTML = `<div class="table-wrap"><table id="staleTable"><thead>${headerTop}${headerStages}</thead><tbody>${body}</tbody>${tfoot}</table></div>`;
  }
  daysInput.addEventListener('change', apply); apply();
  $('#exportStale').onclick = ()=> exportTableXLS('#staleTable','stale_over_'+($('#staleDays').value||'200')+'d.xls');
}

// ====== Tasks: robust import + filters ======
function t_parseDate(s){
  if(!s) return null; const str=String(s).trim();
  let m = str.match(/^(\d{2})\.(\d{2})\.(\d{4})(?:\s+(\d{2}):(\d{2})(?::(\d{2}))?)?$/);
  if(m){ return new Date(+m[3],+m[2]-1,+m[1], +(m[4]||0), +(m[5]||0), +(m[6]||0)) }
  m = str.match(/^(\d{4})-(\d{2})-(\d{2})(?:[ T](\d{2}):(\d{2})(?::(\d{2}))?)?$/);
  if(m){ return new Date(+m[1],+m[2]-1,+m[3], +(m[4]||0), +(m[5]||0), +(m[6]||0)) }
  const ts = Date.parse(str); return isNaN(ts)?null:new Date(ts);
}
function t_filterRows(rows){
  const fc=$('#tFromC')?.value?new Date($('#tFromC').value):null; const tc=$('#tToC')?.value?new Date($('#tToC').value):null;
  const fz=$('#tFromZ')?.value?new Date($('#tFromZ').value):null; const tz=$('#tToZ')?.value?new Date($('#tToZ').value):null;
  const cr=$('#tCreator')?.value||''; const as=$('#tAssignee')?.value||''; const st=$('#tStatus')?.value||'';
  return rows.filter(r=>{ const dc=t_parseDate(r['Дата создания']); const dz=t_parseDate(r['Дата закрытия']);
    if(fc && (!dc || dc<fc)) return false; if(tc && (!dc || dc>tc)) return false;
    if(fz || tz){ if(!dz) return false; if(fz && dz<fz) return false; if(tz && dz>tz) return false; }
    if(cr && (r['Постановщик']||'')!==cr) return false; if(as && (r['Исполнитель']||'')!==as) return false; if(st && (r['Статус']||'')!==st) return false; return true; });
}
function uniqById(rows){const m=new Map();rows.forEach(r=>{const id=r['ID']!=null&&r['ID']!==''?String(r['ID']):null;if(!id)return;if(!m.has(id))m.set(id,r)});return Array.from(m.values())}
function tableKV(a,b,obj){const rows=Object.entries(obj).sort((x,y)=>y[1]-x[1]);return `<table><thead><tr><th>${a}</th><th>${b}</th></tr></thead><tbody>${rows.map(r=>`<tr><td>${r[0]}</td><td><b>${r[1]}</b></td></tr>`).join('')}</tbody></table>`}
function t_populateFilters(src){
  const creators=['',...new Set(src.map(r=>r['Постановщик']||'').filter(Boolean))].sort();
  const assignees=['',...new Set(src.map(r=>r['Исполнитель']||'').filter(Boolean))].sort();
  const statuses=['',...new Set(src.map(r=>r['Статус']||'').filter(Boolean))].sort();
  $('#tCreator').innerHTML = creators.map(v=>`<option value="${v}">${v||'Все постановщики'}</option>`).join('');
  $('#tAssignee').innerHTML = assignees.map(v=>`<option value="${v}">${v||'Все исполнители'}</option>`).join('');
  $('#tStatus').innerHTML = statuses.map(v=>`<option value="${v}">${v||'Все статусы'}</option>`).join('');
}
function t_build(){
  const src=uniqById(t_filterRows(window.TASKS||[]));
  const self=src.filter(r=> (r['Постановщик']&&r['Исполнитель']&&r['Постановщик']===r['Исполнитель']));
  const bySelf=self.reduce((a,r)=>{const k=r['Исполнитель']||'—';a[k]=(a[k]||0)+1;return a},{}); $('#tTbl1').innerHTML=tableKV('Сотрудник','Кол-во задач',bySelf);
  const byCr=src.reduce((a,r)=>{const k=r['Постановщик']||'—';a[k]=(a[k]||0)+1;return a},{}); $('#tTbl2').innerHTML=tableKV('Постановщик','Кол-во задач',byCr);
  const byAs=src.reduce((a,r)=>{const k=r['Исполнитель']||'—';a[k]=(a[k]||0)+1;return a},{}); $('#tTbl3').innerHTML=tableKV('Исполнитель','Кол-во задач',byAs);
  const statuses=[...new Set(src.map(r=>r['Статус']||'—'))].sort(); const assignees=[...new Set(src.map(r=>r['Исполнитель']||'—'))].sort();
  let head='<tr><th>Исполнитель</th>'+statuses.map(s=>`<th>${s}</th>`).join('')+'<th>Итого</th></tr>';
  let body=assignees.map(p=>{ let sum=0; const cells=statuses.map(s=>{ const v=src.filter(r=>(r['Исполнитель']||'—')===p&&(r['Статус']||'—')===s).length; sum+=v; return `<td>${v?'<b>'+v+'</b>':''}</td>`}).join(''); return `<tr><th>${p}</th>${cells}<th>${sum}</th></tr>`}).join('');
  $('#tTbl4').innerHTML=`<table id="tasksStatus"><thead>${head}</thead><tbody>${body}</tbody></table>`;
}

// ====== Import & UI wiring ======
function blobFromText(id,text){ const blob=new Blob([text],{type:'text/csv;charset=utf-8'}); const url=URL.createObjectURL(blob); const a=document.getElementById(id); a.href=url; }
async function importDealsFromFile(file){
  try{ setPill('#dInfoH','Загрузка…',''); const text=await readFileSmart(file); const rows=parseCSVText(text); const {rows:Norm,info}=normalizeDeals(rows); const mode=$('#dModeH').value||'replace'; if(mode==='replace'){APP.DealsA.rows=Norm}else{APP.DealsA.rows=mergeDeals(APP.DealsA.rows||[],Norm)}; APP.lastDealsImport={file:file.name,...info,total:APP.DealsA.rows.length,when:new Date().toISOString()}; setPill('#dInfoH',`✓ ${Norm.length} | всего ${APP.lastDealsImport.total} | пропущено ${info.ignored}`,'ok'); log({file:file.name,...info},'Deals import'); toast('Сделки импортированы и сохранены в списке файлов','ok'); addDealFile(file.name, Norm); buildAll(); }catch(e){console.error(e); setPill('#dInfoH','Ошибка импорта','err'); toast('Ошибка импорта сделок','err'); log(String(e),'Deals import error')} showZeroStates();
}
async function importTasksFromFile(file){
  try{ setPill('#tInfoH','Загрузка…',''); const text=await readFileSmart(file); const rows=parseCSVText(text);
    const normTasks=(rows)=>{ const norm=s=>String(s||'').toLowerCase().replace(/\s+/g,' ').trim(); const cols={}; Object.keys(rows[0]||{}).forEach(c=>cols[c]=norm(c)); const pick=a=>{for(const[orig,n]of Object.entries(cols)){if(a.some(x=>n.includes(x)))return orig}return null}; const cId=pick(['id','ид','номер','task id','key']); const cCr=pick(['постановщик','creator','created by','автор']); const cAs=pick(['исполнитель','assignee','ответственный']); const cSt=pick(['статус','status']); const cTi=pick(['название','тема','задача','title','summary']); const cCrD=pick(['дата создания','создан','created']); const cClD=pick(['дата закрытия','закрыт','closed','completed']); return rows.map(r=>({'ID':r[cId]||r['ID']||'','Название':r[cTi]||r['Название']||'','Постановщик':r[cCr]||r['Постановщик']||'','Исполнитель':r[cAs]||r['Исполнитель']||'','Статус':r[cSt]||r['Статус']||'','Дата создания':r[cCrD]||r['Дата создания']||'','Дата закрытия':r[cClD]||r['Дата закрытия']||''})) };
    window.TASKS = normTasks(rows); setPill('#tInfoH',`✓ ${window.TASKS.length}`,'ok'); t_populateFilters(window.TASKS); t_build(); toast('Задачи импортированы','ok'); }catch(e){console.error(e); setPill('#tInfoH','Ошибка импорта','err'); toast('Ошибка импорта задач','err')}
}

function buildAll(){ const tab=document.querySelector('.tab.active')?.dataset.tab||'mismatch'; if(tab==='mismatch')buildMismatch(); if(tab==='dashboard')buildCharts(); if(tab==='stale')buildStale(); if(tab==='tasks')t_build(); if(tab==='compare')refreshFileLists(); }

function switchTab(id){ document.querySelectorAll('.page').forEach(p=>p.style.display='none'); document.getElementById(id).style.display='grid'; document.querySelectorAll('.tab').forEach(t=>t.classList.remove('active')); document.querySelector(`.tab[data-tab="${id}"]`).classList.add('active'); buildAll() }

document.addEventListener('DOMContentLoaded', ()=>{
  document.querySelectorAll('.tab').forEach(t=> t.addEventListener('click', ()=> switchTab(t.dataset.tab)));
  ['#tFromC','#tToC','#tFromZ','#tToZ'].forEach(sel=>{ const el=$(sel); if(el) el.setAttribute('type','date'); });
  blobFromText("dlDealsTpl", "ID,\u0422\u0438\u043f,\u0421\u0442\u0430\u0434\u0438\u044f \u0441\u0434\u0435\u043b\u043a\u0438,\u0414\u0430\u0442\u0430 \u0438\u0437\u043c\u0435\u043d\u0435\u043d\u0438\u044f,\u041e\u0442\u0432\u0435\u0442\u0441\u0442\u0432\u0435\u043d\u043d\u044b\u0439,\u0414\u0430\u0442\u0430 \u0441\u043e\u0437\u0434\u0430\u043d\u0438\u044f,\u041c\u041f\u041e \u041e\u0442\u0432\u0435\u0442\u0441\u0442\u0432\u0435\u043d\u043d\u044b\u0439,\u0412\u043e\u0440\u043e\u043d\u043a\u0430,\u041c\u0435\u043d\u0435\u0434\u0436\u0435\u0440 \u041e\u041f,\u041d\u0430\u0437\u0432\u0430\u043d\u0438\u0435 \u0441\u0434\u0435\u043b\u043a\u0438,\u041a\u0435\u043c \u0438\u0437\u043c\u0435\u043d\u0435\u043d\u0430,\n");
  blobFromText("dlTasksTpl", "ID,\u041d\u0430\u0437\u0432\u0430\u043d\u0438\u0435,\u041a\u0440\u0430\u0439\u043d\u0438\u0439 \u0441\u0440\u043e\u043a,\u041f\u043e\u0441\u0442\u0430\u043d\u043e\u0432\u0449\u0438\u043a,\u0418\u0441\u043f\u043e\u043b\u043d\u0438\u0442\u0435\u043b\u044c,\u0421\u0442\u0430\u0442\u0443\u0441,\u041f\u0440\u043e\u0435\u043a\u0442,\u0414\u0430\u0442\u0430 \u0441\u043e\u0437\u0434\u0430\u043d\u0438\u044f,\u0414\u0430\u0442\u0430 \u0438\u0437\u043c\u0435\u043d\u0435\u043d\u0438\u044f,\u0414\u0430\u0442\u0430 \u0437\u0430\u043a\u0440\u044b\u0442\u0438\u044f,\u041c\u0435\u043d\u0435\u0434\u0436\u0435\u0440 \u041e\u041f,\u0420\u0435\u0433\u0438\u043e\u043d\n");
  // (placeholders replaced by Python below)
  // Auto-import
  $('#dFileH').addEventListener('change', e=>{ const f=e.target.files[0]; if(f) importDealsFromFile(f) });
  $('#tFileH').addEventListener('change', e=>{ const f=e.target.files[0]; if(f) importTasksFromFile(f) });
  // Exports
  $('#exportDash').addEventListener('click', exportDashboard);
  $('#exportTasksSummary').addEventListener('click', ()=> exportMultipleTablesXLS(['#tTbl1 table','#tTbl2 table','#tTbl3 table','#tTbl4 table'],'tasks_summary.xls'));
  // Compare (files)
  $('#fileList').addEventListener('change', fillMetaFromSelected);
  $('#fileRename').addEventListener('click', renameSelected);
  $('#fileDelete').addEventListener('click', deleteSelected);
  $('#metaAuto').addEventListener('click', ()=> updateSelectedMeta(true));
  $('#metaSave').addEventListener('click', ()=> updateSelectedMeta(false));
  $('#compareMulti').addEventListener('click', compareMulti);
  $('#compareAB').addEventListener('click', compareAB);
  $('#toggleLog').addEventListener('click', ()=>{ const el=$('#logPanel'); el.style.display = el.style.display==='none'?'block':'none' });
  switchTab('compare');
});
