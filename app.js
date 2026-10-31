(function(){
"use strict";

/* ---------- small utils ---------- */
const $ = (s,el)=> (el||document).querySelector(s);
const $$ = (s,el)=> Array.from((el||document).querySelectorAll(s));
const todayISO = ()=>{ const d = new Date(); return d.getFullYear()+'-'+String(d.getMonth()+1).padStart(2,'0')+'-'+String(d.getDate()).padStart(2,'0'); };
const WEEKDAYS = ['Sunday','Monday','Tuesday','Wednesday','Thursday','Friday','Saturday'];

function safeLS(fn, fallback){ try{ return fn(); }catch(e){ return fallback; } }
function lsGet(k,d){ const S = window.KR_STORE || {}; return (k in S && S[k]!=null) ? JSON.parse(JSON.stringify(S[k])) : d; }
function lsSet(k,v){ const S = window.KR_STORE; if(!S) return; S[k] = JSON.parse(JSON.stringify(v==null?null:v)); if(window.KR_SAVE) window.KR_SAVE(); }

function normHeader(h){ return String(h||'').toLowerCase().replace(/[^a-z0-9]/g,''); }

const HEADER_MAP = [
  ['list','list'],['status','status'],['box','newbox'],['customername','name'],
  ['packageduration','packageduration'],['mealplan','mealplan'],['allerg','allergies'],
  ['dislike','dislike'],['customernotes','custnotes'],['packnote','packnote'],
  ['joiningdate','joindate'],['duedate','duedate'],['offday','offdays'],
  ['note','note']
];
function mapHeaderByName(h){
  const n = normHeader(h);
  const lm = /^location([ab])(zone|street|building|fromday|today|fromdate|todate)$/.exec(n);
  if(lm) return 'loc'+lm[1]+lm[2];
  if(n==='planapp') return 'planapp';
  if(n.includes('contact')) return 'contact';
  if(n.includes('remaining')) return 'remaining';
  if(n==='today') return 'today';
  if(n.includes('freezeonoff')) return 'freezeonoff';
  if(n.includes('freezeperiods')) return 'freezeperiods';
  if(n.includes('frozendays')) return 'frozendays';
  if(n.includes('duedatebefore')) return 'duebase';
  if(n.includes('freezedate')) return 'freezedate';
  if(n.includes('offreezedays') || n.includes('freezedays')) return 'freezedays';
  if(n.includes('box')) return 'newbox';
  if(n.includes('customername')) return 'name';
  if(n==='name') return 'name';
  if(n.includes('packageduration')) return 'packageduration';
  if(n.includes('mealplan')) return 'mealplan';
  if(n.includes('allerg') || n.includes('allegi')) return 'allergies';
  if(n.includes('dislikeremov') || n.includes('dislikeingredient')) return 'removables';
  if(n.includes('dislikemeal')) return 'dislikemeals';
  if(n.includes('dislike')) return 'dislike';
  if(n.includes('customernotes')) return 'custnotes';
  if(n.includes('packnote')) return 'packnote';
  if(n.includes('joiningdate')) return 'joindate';
  if(n.includes('duedate')) return 'duedate';
  if(n.includes('offday')) return 'offdays';
  if(n==='list' || n==='hk') return 'list';
  if(n==='status') return 'status';
  if(n.includes('note')) return 'note';
  return null;
}
// Stable column orders observed in real exports — used as a fallback when a
// header cell's text is corrupted (e.g. "LIST" read back as "3013 0133").
const POSITIONAL_OLD = ['list','status',null,'newbox','name','packageduration','mealplan','allergies','dislike','custnotes','packnote','joindate',null,'duedate',null,'offdays',null,null,'note'];
const POSITIONAL_NEW = ['list','status',null,'newbox','name','packageduration','mealplan','allergies','removables','dislikemeals','custnotes','packnote','joindate',null,'duedate',null,'offdays',null,null,'note'];
function mapHeader(h, colIndex, newLayout){
  const byName = mapHeaderByName(h);
  if(byName) return byName;
  const pos = newLayout ? POSITIONAL_NEW : POSITIONAL_OLD;
  if(colIndex!=null && colIndex < pos.length) return pos[colIndex];
  return null;
}

function excelDateToISO(v){
  if(v==null || v==='') return null;
  if(v instanceof Date && !isNaN(v)) return v.toISOString().slice(0,10);
  if(typeof v === 'number'){
    const d = XLSX.SSF.parse_date_code(v);
    if(!d) return null;
    const dt = new Date(Date.UTC(d.y, d.m-1, d.d));
    return dt.toISOString().slice(0,10);
  }
  const s = String(v).trim();
  if(!s) return null;
  const d2 = new Date(s);
  if(!isNaN(d2)) return d2.toISOString().slice(0,10);
  return null;
}

/* ---------- state ---------- */
let roster = [];                           // the customer data sheet: one entry per customer
let menu = lsGet('kr_menu', null);
let overrides = lsGet('kr_overrides', {}); // box -> {grams, note}
let runResults = null;
const str = v=> String(v==null ? '' : v).trim();
const uid = ()=> Date.now().toString(36)+Math.random().toString(36).slice(2,7);

const PLANS = [
  ['3 meals + 2 snacks','3 meals + 2 snacks (BF+L+D)'],
  ['3 meals + 2 snacks (no breakfast)','3 meals + 2 snacks (no breakfast)'],
  ['3 meals + 1 snack','3 meals + 1 snack (BF+L+D)'],
  ['3 meals + 1 snack (no breakfast)','3 meals + 1 snack (no breakfast)'],
  ['3 meals','3 meals (BF+L+D) no snack'],
  ['3 meals (no breakfast)','3 meals (no breakfast) no snack'],
  ['2 meals + 2 snacks','2 meals + 2 snacks'],
  ['2 meals + 2 snacks (breakfast and lunch)','2 meals + 2 snacks (breakfast and lunch)'],
  ['2 meals + 1 snack','2 meals + 1 snack'],
  ['2 meals + 1 snack (breakfast and lunch)','2 meals + 1 snack (breakfast and lunch)'],
  ['2 meals','2 meals no snack'],
  ['1 meal + 1 snack (business plan)','1 meal + 1 snack (business plan)']
];
const PKG_LABEL = {keto:'Keto', macro:'Macro', simple:'Simple Diet', app:'App'};
const PKG_LIST = {keto:'KETO', macro:'MACRO', simple:'SIMPLE DIET', app:'APP'};
const blankLoc = ()=>({zone:'',street:'',building:'',fromDay:'',toDay:'',fromDate:'',toDate:''});

/* ---------- dates, off days, freeze, due date, remaining days ----------
   Off days and frozen days do not use up the package: the due date is the day the last meal of the package is served,
   and every frozen serving day pushes it one serving day further. */
function addDaysISO(iso, n){ const d = new Date(iso+'T00:00:00Z'); d.setUTCDate(d.getUTCDate()+n); return d.toISOString().slice(0,10); }
function weekdayOfISO(iso){ return WEEKDAYS[new Date(iso+'T00:00:00Z').getUTCDay()]; }
function offSetOf(e){ const s = new Set(e.daysOff||[]); if(e.hk==='simple') s.add('Friday'); return s; }
function servingDay(e, iso){ return !offSetOf(e).has(weekdayOfISO(iso)); }
function isFrozenNow(e){ return (e.freezePeriods||[]).some(p=>p.to==null); }
function frozenOn(c, iso){ return (c.freezePeriods||[]).some(p=> p.from<=iso && (p.to==null || iso<=p.to)); }
function frozenServingDays(e, today){
  let n = 0; const seen = new Set();
  (e.freezePeriods||[]).forEach(p=>{
    const to = p.to==null ? (today>p.from ? today : p.from) : p.to;
    let d = p.from, g = 0;
    while(d<=to && g++<800){ if(!seen.has(d) && servingDay(e,d)){ seen.add(d); n++; } d = addDaysISO(d,1); }
  });
  return n;
}
function advanceServing(e, base, n){ let d = base, g = 0; while(n>0 && g++<2000){ d = addDaysISO(d,1); if(servingDay(e,d)) n--; } return d; }
function computeDueBase(e){
  if(!e.start || !(+e.duration>0)) return '';
  let d = e.start, left = +e.duration, g = 0;
  while(g++<2000){ if(servingDay(e,d)){ left--; if(left===0) return d; } d = addDaysISO(d,1); }
  return '';
}
function dueEff(e, today){ return e.dueBase ? advanceServing(e, e.dueBase, frozenServingDays(e, today)) : ''; }
function remainingDays(e, today){
  const due = dueEff(e, today); if(!due) return '';
  let n = 0, d = today, g = 0;
  while(d<due && g++<2000){ d = addDaysISO(d,1); if(servingDay(e,d)) n++; }
  return n;
}
function setFreeze(e, on, today){
  e.freezePeriods = e.freezePeriods || [];
  const open = e.freezePeriods.find(p=>p.to==null);
  if(on && !open) e.freezePeriods.push({from:today, to:null});
  if(!on && open){
    const y = addDaysISO(today,-1);
    if(y < open.from) e.freezePeriods = e.freezePeriods.filter(p=>p!==open); else open.to = y;
  }
}
function statusInfo(e, today){
  const raw = str(e.status).toUpperCase() || 'ACTIVE';
  if(/EXPIR/.test(raw)) return ['Expired','pill-bad'];
  if(isFrozenNow(e) || /FREEZ|FROZEN/.test(raw)) return ['Frozen','pill-warn'];
  if(e.start && today < e.start) return ['Starts '+e.start,'pill-dim'];
  const due = dueEff(e, today); if(due && today > due) return ['Ended','pill-dim'];
  if(raw!=='ACTIVE') return [raw,'pill-dim'];
  return ['Active','pill-ok'];
}

/* ---------- a sheet entry becomes the row shape the engine reads ---------- */
function customNotes(e){
  const g = e.grams || {};
  const part = (lab, p, c)=> (p||c) ? lab+': '+(p?p+'G PRO':'')+(p&&c?' + ':'')+(c?c+'G CARB':'') : '';
  const lines = [];
  if(e.gramType==='custom'){ lines.push(part('LUNCH', g.lp, g.lc)); lines.push(part('DINNER', g.dp, g.dc)); }
  if(e.notes) lines.push(e.notes);
  return lines.filter(Boolean).join('\n');
}
function variantWord(e){ return (e.hk==='keto'||e.hk==='app') ? '' : (e.gramType==='hp' ? 'HP ' : 'Classic '); }
function displayPlan(e){ return e.plan ? PKG_LABEL[e.hk]+' '+variantWord(e)+e.plan : (e.mealplan||''); }
function enginePlan(e){
  if(!e.plan) return e.mealplan||'';
  const p = PLANS.find(x=>x[0]===e.plan);
  return PKG_LABEL[e.hk]+' '+variantWord(e)+(p?p[1]:e.plan);
}
function rosterRow(e){
  const today = todayISO(), raw = str(e.status).toUpperCase() || 'ACTIVE';
  return {
    id:e.id, list:PKG_LIST[e.hk]||'APP', status: /EXPIR/.test(raw) ? 'EXPIRED' : (raw==='ACTIVE' ? 'ACTIVE' : raw),
    contact:e.contact, newbox:e.box, name:e.name, mealplan:enginePlan(e), allergies:e.allergies||'', removables:e.removables||'',
    dislikemeals:e.dislikemeals||'', custnotes:customNotes(e), packnote:e.packnote||'', packageduration:e.duration,
    joindate:e.start||null, duedate:dueEff(e,today)||null, offdays:(e.daysOff||[]).join(', '), freezePeriods:e.freezePeriods||[], enrolled:true
  };
}
function allCustomers(){ return roster.map(rosterRow); }

/* ---------- loading a customer data workbook ---------- */
function parseDaysOff(t){
  const s = str(t).toUpperCase(), out = [];
  WEEKDAYS.forEach(d=>{ if(new RegExp('\\b'+d.slice(0,3).toUpperCase()).test(s)) out.push(d); });
  return out;
}
function parsePeriods(t){
  return str(t).split(';').map(s=>s.trim()).filter(Boolean).map(s=>{ const a = s.split('>'); return {from:(a[0]||'').trim(), to:(a[1]||'').trim()||null}; })
    .filter(p=>/^\d{4}-\d\d-\d\d$/.test(p.from));
}
const periodsText = e=> (e.freezePeriods||[]).map(p=>p.from+'>'+(p.to||'')).join('; ');
function entryFromRec(rec, mine){
  const hk = normDiet(rec.list) || 'app';
  let status = str(rec.status).toUpperCase() || 'ACTIVE';
  const mp = str(rec.mealplan);
  const lab = PLANS.find(p=>p[0]===str(rec.planapp));
  const dur = parseInt(rec.packageduration,10);
  const periods = parsePeriods(rec.freezeperiods);
  if(mine && periods.length && /FROZEN/.test(status)) status = 'ACTIVE';
  const loc = k=>({zone:str(rec['loc'+k+'zone']), street:str(rec['loc'+k+'street']), building:str(rec['loc'+k+'building']),
    fromDay:str(rec['loc'+k+'fromday']), toDay:str(rec['loc'+k+'today']), fromDate:rec['loc'+k+'fromdate']||'', toDate:rec['loc'+k+'todate']||''});
  const e = {
    id:uid(), hk, status, contact:str(rec.contact), box:str(rec.newbox).replace(/\s+/g,'').toUpperCase(), name:str(rec.name),
    plan: lab ? lab[0] : '', mealplan:mp, gramType: classifyVariant(mp)==='hp' ? 'hp' : 'classic', grams:{lp:'',lc:'',dp:'',dc:''},
    allergies:str(rec.allergies), removables:str(rec.removables||rec.dislike), dislikemeals:str(rec.dislikemeals),
    notes:str(rec.custnotes), packnote:str(rec.packnote), duration: isNaN(dur) ? '' : dur,
    start:rec.joindate||'', dueBase:rec.duebase||rec.duedate||'', daysOff:parseDaysOff(rec.offdays),
    freezePeriods:periods, locA:loc('a'), locB:loc('b'),
    xNote:str(rec.note), xFreeze:[str(rec.freezedate), str(rec.freezedays)].filter(Boolean).join(' | ')
  };
  e.dueAuto = !!(e.duration && e.start && computeDueBase(e)===e.dueBase);
  return e;
}
function migrateEnrolled(o){
  const L = o.loc1||{}, M = o.loc2||{};
  const loc = x=>Object.assign(blankLoc(), {zone:x.zone||'', street:x.street||'', building:x.building||''});
  return {id:o.id||uid(), hk:o.package, status:o.expired?'EXPIRED':'ACTIVE', contact:o.contact||'', box:o.box||'', name:o.name||'',
    plan:o.plan||'', mealplan:'', gramType:o.gramType||'classic', grams:o.grams||{lp:'',lc:'',dp:'',dc:''}, allergies:o.allergies||'',
    removables:o.removables||'', dislikemeals:o.dislikemeals||'', notes:o.notes||'', packnote:o.foodSchedule?'FOOD SCHEDULE':'',
    duration:'', start:o.start||'', dueBase:o.end||'', dueAuto:false, daysOff:o.daysOff||[],
    freezePeriods:(o.freezeDates||[]).map(d=>({from:d,to:d})), locA:loc(L), locB:loc(M), xNote:'', xFreeze:''};
}
function initRoster(){
  const saved = lsGet('kr_roster', null);
  if(saved){ roster = saved; return; }
  roster = [];
  lsGet('kr_customers', []).forEach(rec=>{ try{ roster.push(entryFromRec(rec,false)); }catch(err){} });
  lsGet('kr_enrolled', []).forEach(o=>{ try{ roster.push(migrateEnrolled(o)); }catch(err){} });
  if(roster.length) lsSet('kr_roster', roster);
}
function saveRoster(){
  lsSet('kr_roster', roster);
  runResults = null;
  refreshEligCount();
  renderCustStatus();
  if(typeof renderGrid==='function') renderGrid();
}
function applyIncoming(list, mode){
  let added = 0, updated = 0;
  if(mode==='replace'){ roster = list; added = list.length; }
  else {
    const used = new Set(), nm = x=>String(x||'').toUpperCase().replace(/\s+/g,' ');
    list.forEach(e=>{
      let i = e.box ? roster.findIndex((o,ix)=>!used.has(ix) && o.box===e.box && nm(o.name)===nm(e.name)) : -1;
      if(i<0 && e.box) i = roster.findIndex((o,ix)=>!used.has(ix) && o.box===e.box);
      if(i<0 && !e.box) i = roster.findIndex((o,ix)=>!used.has(ix) && !o.box && nm(o.name)===nm(e.name));
      if(i>=0){ e.id = roster[i].id; roster[i] = e; used.add(i); updated++; }
      else { roster.push(e); used.add(roster.length-1); added++; }
    });
  }
  saveRoster();
  return {added, updated};
}
function renderCustStatus(msg){
  const el = $('#custStatus'); if(!el) return;
  const n = $('#rosterCount'); if(n) n.textContent = String(roster.length);
  if(msg){ el.innerHTML = msg; return; }
  el.innerHTML = roster.length ? `<span class="status-pill pill-ok">${roster.length} customers</span>` : 'No customers yet. Register one or load a workbook.';
}

$('#custFile').addEventListener('change', (e)=>{
  const f = e.target.files[0];
  if(!f) return;
  const reader = new FileReader();
  reader.onload = (ev)=>{
    try{
      const wb = XLSX.read(ev.target.result, {type:'binary', cellDates:true});
      const sheetName = wb.SheetNames.find(n=>/customer\s*data/i.test(n)) || wb.SheetNames.find(n=>/active/i.test(n)) || wb.SheetNames[0];
      const ws = wb.Sheets[sheetName];
      const rows = XLSX.utils.sheet_to_json(ws, {header:1, raw:true, defval:''});
      if(!rows.length) throw new Error('empty sheet');
      const headerRow = rows[0];
      const mine = headerRow.some(x=>normHeader(x)==='planapp');
      const newLayout = headerRow.some(x=>/removable/i.test(String(x)));
      const keyFor = headerRow.map((x,idx)=> mine ? mapHeaderByName(x) : mapHeader(x, idx, newLayout));
      const incoming = [];
      for(let i=1;i<rows.length;i++){
        const r = rows[i];
        if(!r || r.every(c=>c==='' || c==null)) continue;
        const rec = {};
        keyFor.forEach((k,idx)=>{ if(k) rec[k] = r[idx]; });
        if(!str(rec.name)) continue;
        ['joindate','duedate','duebase','locafromdate','locatodate','locbfromdate','locbtodate'].forEach(k=>{ if(k in rec) rec[k] = excelDateToISO(rec[k]); });
        incoming.push(entryFromRec(rec, mine));
      }
      if(!incoming.length) throw new Error('no customer rows found');
      const res = applyIncoming(incoming, $('#loadMode').value);
      renderCustStatus(`<span class="status-pill pill-ok">${incoming.length} customers read</span> &nbsp; <span class="foot-note">${res.added} added, ${res.updated} updated from “${escHtml(sheetName)}”. ${roster.length} in the sheet now.</span>`);
    }catch(err){
      renderCustStatus(`<span class="status-pill pill-bad">Couldn't read that file</span> ${escHtml(String(err.message||err))}`);
    }
  };
  reader.readAsBinaryString(f);
});

/* ---------- tabs and dates ---------- */
const prodDateInput = $('#prodDate');
{ let sd = lsGet('kr_date',''); if(!sd || sd < todayISO()) sd = todayISO(); prodDateInput.value = sd; }
let currentTab = 'home';
const RUN_TABS = ['sheet','manual','orders','detail'];
function showTab(name){
  currentTab = name;
  $$('.tab').forEach(b=>b.setAttribute('aria-selected', b.dataset.tab===name ? 'true' : 'false'));
  $$('.panel').forEach(p=>p.classList.remove('active'));
  $('#panel-'+name).classList.add('active');
  if(RUN_TABS.includes(name)) runEngine();
}
$$('.tab').forEach(btn=> btn.addEventListener('click', ()=> showTab(btn.dataset.tab)));
function refreshDateReadout(){
  const d = new Date(prodDateInput.value+'T00:00:00');
  $('#dateReadout').textContent = isNaN(d) ? '—' : ('Production · '+WEEKDAYS[d.getDay()]);
  const t = new Date(todayISO()+'T00:00:00');
  $('#todayReadout').textContent = 'Today · '+WEEKDAYS[t.getDay()]+' '+todayISO();
}
prodDateInput.addEventListener('change', ()=>{
  lsSet('kr_date', prodDateInput.value);
  refreshDateReadout();
  refreshEligCount();
  if($('#autoFollowDate') && $('#autoFollowDate').checked) autofillMenu(prodDateInput.value);
  runResults = null;
  if(RUN_TABS.includes(currentTab)) runEngine();
});

/* ---------- menu builder ----------
   Shape (matches the monthly rotation docs):
   menu.simple[slot] = {main,opt}                       — all 5 slots
   menu.macro[slot]  = {main,opt}                        — breakfast/amsnack/pmsnack
   menu.macro.lunch  = {main, opt:{classic,hp}}           — main shared, option splits by variant
   menu.macro.dinner = {main, opt:{classic,hp}}
   menu.keto[slot]   = {main,opt}                        — all 5 slots
*/
const DIET_LINES = [
  {key:'simple', label:'Simple Diet (off Fridays)', slots:['breakfast','lunch','dinner','amsnack','pmsnack']},
  {key:'macro',  label:'Macro — Classic & High Protein', slots:['breakfast','lunch','dinner','amsnack','pmsnack']},
  {key:'keto',   label:'Keto', slots:['breakfast','lunch','dinner','amsnack','pmsnack']}
];
const SLOT_LABEL = {breakfast:'Breakfast', lunch:'Lunch', dinner:'Dinner', amsnack:'AM Snack', pmsnack:'PM Snack', main3:'Main meal #3', extra:'Special add-ons'};
const SPLIT_OPT_SLOTS = {macro:['lunch','dinner']};

function emptyMenu(){
  const m = {};
  DIET_LINES.forEach(line=>{
    m[line.key] = {};
    line.slots.forEach(slot=>{
      const splitHere = (SPLIT_OPT_SLOTS[line.key]||[]).includes(slot);
      m[line.key][slot] = splitHere ? {main:'', opt:{classic:'', hp:''}} : {main:'', opt:''};
    });
  });
  return m;
}
if(!menu) menu = emptyMenu();

let menuData = null; // loaded from menu_data.json (the monthly rotation)
const JSON_SLOT_KEY = {breakfast:'breakfast', lunch:'lunch', dinner:'dinner', amsnack:'amsnack', pmsnack:'pmsnack'};

async function loadMenuData(){
  try{
    const res = await fetch('menu_data.json');
    if(!res.ok) throw new Error('fetch failed');
    menuData = await res.json();
    $('#menuSourceNote').innerHTML = `Monthly rotation loaded (Simple Diet, Macro, Keto). <strong>Auto-fill</strong> pulls today's dishes from it — you can still edit any box by hand afterward.`;
  }catch(e){
    menuData = null;
    $('#menuSourceNote').textContent = "Couldn't load the monthly rotation file — fill the menu in by hand below.";
  }
}

function daysBetween(aISO, bISO){
  return Math.round((new Date(bISO+'T00:00:00') - new Date(aISO+'T00:00:00')) / 86400000);
}
function resolveMenuForDate(lineKey, dateISO){
  if(!menuData || !menuData[lineKey]) return null;
  const ld = menuData[lineKey];
  const weekday = weekdayFor(dateISO);
  if(ld.offDay && weekday === ld.offDay) return null;
  const diff = daysBetween(ld.anchor, dateISO);
  const n = ld.weeks.length;
  const weekIdx = (((Math.floor(diff/7)) % n) + n) % n;
  const dayData = ld.weeks[weekIdx].days[weekday];
  return dayData || null;
}

function autofillMenu(dateISO){
  const fresh = emptyMenu();
  DIET_LINES.forEach(line=>{
    const dayData = resolveMenuForDate(line.key, dateISO);
    if(!dayData) return; // off day or no data — leave blank
    line.slots.forEach(slot=>{
      const jk = JSON_SLOT_KEY[slot];
      const mainVal = dayData[jk] || '';
      const optRaw = dayData[jk+'_opt'];
      const splitHere = (SPLIT_OPT_SLOTS[line.key]||[]).includes(slot);
      if(splitHere){
        if(optRaw && typeof optRaw === 'object'){
          fresh[line.key][slot] = {main:mainVal, opt:{classic:optRaw.classic||'', hp:optRaw.hp||''}};
        } else {
          fresh[line.key][slot] = {main:mainVal, opt:{classic:optRaw||'', hp:optRaw||''}};
        }
      } else {
        const optVal = (optRaw && typeof optRaw === 'object') ? (optRaw.classic || optRaw.hp || '') : (optRaw || '');
        fresh[line.key][slot] = {main:mainVal, opt:optVal};
      }
    });
  });
  menu = fresh;
  lsSet('kr_menu', menu);
  renderMenuBuilder();
}

const SAMPLE_MENU = {
  simple:{breakfast:{main:'Egg w/ Mushroom',opt:'Croissant w/ Turkey'}, lunch:{main:'Pasta Bolognese',opt:'Asian Grilled Chicken'}, dinner:{main:'Chicken Fajita Wrap',opt:'Chicken Tikka w/ Bukhari Rice'}, amsnack:{main:'Vanilla Muffins',opt:'Milk Cake'}, pmsnack:{main:'Green Sauce Salad',opt:'Chicken Madrouba'}},
  macro:{breakfast:{main:'Waffle w/ Maple & Berries',opt:'Egg w/ Mushroom'},
    lunch:{main:'Beef Bolognese', opt:{classic:'Chicken Tikka', hp:'Salmon Pestochurri Pasta'}},
    dinner:{main:'Chicken Fajita w/ Rice', opt:{classic:'Asian Grilled Chicken w/ Potato', hp:'Veggie Chicken Salad'}},
    amsnack:{main:'Vanilla Muffins',opt:'Pistachio Cake'}, pmsnack:{main:'Halloumi Salad',opt:'Chicken Madrouba'}},
  keto:{breakfast:{main:'Keto Waffles w/ Maple & Berries',opt:'Keto Baked Cauliflower w/ Buffalo Sauce'}, lunch:{main:'Keto Chicken Curry w/ Keto Bread',opt:'Beef w/ Tomato Salsa'}, dinner:{main:'Keto BBQ Chicken Pizza',opt:'Keto Chicken Eggplant Roll'}, amsnack:{main:'Keto Green Sauce Salad',opt:'Keto Caesar Salad'}, pmsnack:{main:'Bounty',opt:'Keto Milk Cake'}}
};

function slotField(path, value, label){
  return `<div class="field"><label>${label}</label><input type="text" class="menu-input" data-path="${path}" value="${(value||'').replace(/"/g,'&quot;')}" placeholder="dish name"></div>`;
}

function renderMenuBuilder(){
  const host = $('#menuBuilder');
  host.innerHTML = DIET_LINES.map(line=>{
    const slotsHtml = line.slots.map(slot=>{
      const splitHere = (SPLIT_OPT_SLOTS[line.key]||[]).includes(slot);
      const v = menu[line.key][slot];
      if(splitHere){
        return `<div class="slot">
          <label>${SLOT_LABEL[slot]} — main dish</label>
          <div class="row">${slotField(`${line.key}.${slot}.main`, v.main,'Main dish')}</div>
          <label style="margin-top:8px;">Option 1</label>
          <div class="row">${slotField(`${line.key}.${slot}.opt.classic`, v.opt.classic,'Option')}</div>
          <label style="margin-top:8px;">Option 2</label>
          <div class="row">${slotField(`${line.key}.${slot}.opt.hp`, v.opt.hp,'Option')}</div>
        </div>`;
      }
      return `<div class="slot">
        <label>${SLOT_LABEL[slot]}</label>
        <div class="row">${slotField(`${line.key}.${slot}.main`, v.main,'Main')}${slotField(`${line.key}.${slot}.opt`, v.opt,'Option')}</div>
      </div>`;
    }).join('');
    return `<div class="card"><h2>${line.label}</h2><div class="menu-grid">${slotsHtml}</div></div>`;
  }).join('');

  $$('.menu-input', host).forEach(inp=>{
    inp.addEventListener('input', ()=>{
      const path = inp.dataset.path.split('.');
      let obj = menu;
      for(let i=0;i<path.length-1;i++) obj = obj[path[i]];
      obj[path[path.length-1]] = inp.value;
      lsSet('kr_menu', menu);
    });
  });
}
renderMenuBuilder();
$('#loadSample').addEventListener('click', ()=>{
  menu = JSON.parse(JSON.stringify(SAMPLE_MENU));
  lsSet('kr_menu', menu);
  renderMenuBuilder();
});
$('#autofillBtn').addEventListener('click', ()=> autofillMenu(prodDateInput.value));
const autoFollowBox = $('#autoFollowDate');
autoFollowBox.checked = lsGet('kr_autofollow', true);
autoFollowBox.addEventListener('change', ()=> lsSet('kr_autofollow', autoFollowBox.checked));

loadMenuData().then(()=>{
  const neverFilled = !lsGet('kr_menu_filled_once', false);
  if(menuData && neverFilled){
    autofillMenu(prodDateInput.value);
    lsSet('kr_menu_filled_once', true);
  }
});

/* ---------- decision engine ---------- */
const upper = s=>String(s||'').toUpperCase();
const SLOT_ORDER = ['breakfast','lunch','dinner','main3','amsnack','pmsnack'];
const LINE_ORDER = ['keto','simple','macro'];
const LINE_LABEL = {keto:'Keto', simple:'Simple Diet', macro:'Macro', app:'App'};

function normDiet(listVal){
  const s = upper(listVal);
  if(s.includes('KETO')) return 'keto';
  if(s.includes('MACRO')) return 'macro';
  if(s.includes('SIMPLE')) return 'simple';
  if(s.includes('APP')) return 'app';
  return null;
}
function statusCheck(c){
  const s = String(c.status||'').trim().toUpperCase();
  if(s.includes('FREEZE') || s.includes('FROZEN')) return {ok:false, reason:'Frozen'};
  if(s.includes('EXPIR')) return {ok:false, reason:'Expired'};
  if(s==='ACTIVE') return {ok:true};
  return {ok:false, reason: s ? ('Status: '+s) : 'Status blank'};
}
function hasFoodSchedule(c){
  const t = String(c.packnote||'').toUpperCase();
  return t.includes('FOOD SCHEDULE') && !/\bNO\s+FOOD\s+SCHEDULE/.test(t);
}
function dateOk(c, prodDateISO){
  if(c.joindate && prodDateISO < c.joindate) return {ok:false, reason:'Before joining date ('+c.joindate+')'};
  if(c.duedate && prodDateISO > c.duedate) return {ok:false, reason:'Past due date ('+c.duedate+')'};
  return {ok:true};
}
function offDayHit(c, weekdayName){
  const t = String(c.offdays||'');
  if(!t) return false;
  return new RegExp(weekdayName, 'i').test(t);
}
function weekdayFor(dateISO){ return WEEKDAYS[new Date(dateISO+'T00:00:00').getDay()]; }

function classifyVariant(mealPlanText){
  const t = upper(mealPlanText);
  if(t.includes('CLASSIC')) return 'classic';
  if(/\bHP\b/.test(t) || t.includes('HIGH PROT')) return 'hp';
  return 'unknown';
}
// "3 meals", "2meals", "2 (BF+L) meal", "2( L+D)  meal", and the shorthand "3M/2S"
function mealCount(t){ const m=/(\d)\s*(?:\([^)]*\)\s*)?(?:MAIN\s*)?(?:MEALS?|M)\b/.exec(upper(t)); return m? +m[1] : null; }
function snackCount(t){
  const u = upper(t);
  if(u.includes('NO SNACK')) return 0;
  const m=/(\d)\s*(?:SNACKS?|S)\b/.exec(u); return m? +m[1] : 2;
}

// Which meals a package contains:
//  3M/2S = BF+L+D+AM+PM | 2M = L+D (or BF+L / BF+D when stated) | 3 main meals (no BF) = L+D+#3
//  3M/1S = BF+L+D + one snack | business (1 meal) = lunch + 1 snack
function composition(planText, extraText){
  const p = upper(planText), x = upper(planText+' '+(extraText||''));
  const n = mealCount(p);
  const snacks = snackCount(p);
  const noBF = /NO\s*BREAK\s*FAST|NO\s*BF\b/.test(x);
  const bfWord = /BREAKFAST|\bBF\b/.test(p.replace(/NO\s*BREAK\s*FAST|NO\s*BF\b/g,''));
  const combo = (a,b)=> new RegExp('(?:'+a+')\\s*(?:AND|\\+|&|,)\\s*(?:'+b+')').test(x);
  if(n===null) return {ok:false, reason:'Meal plan not recognised ("'+String(planText||'').trim()+'")'};
  let slots, kind='normal';
  if(n===1){ kind='business'; slots=[ bfWord ? 'breakfast' : ((/DINNER/.test(p)&&!/LUNCH/.test(p)) ? 'dinner' : 'lunch') ]; }
  else if(n===2){
    if(combo('BF|BREAKFAST','\\bL\\b|LUNCH')) slots=['breakfast','lunch'];
    else if(combo('BF|BREAKFAST','\\bD\\b|DINNER')) slots=['breakfast','dinner'];
    else slots=['lunch','dinner'];
  } else if(n===3){
    const main3 = noBF || /MAIN\s*MEALS?/.test(p) || /\(L\s*(?:\+|&|AND)\s*D\)|LUNCH\s*(?:\+|&|AND)\s*DINNER/.test(p) || (/NO SNACK|ONLY/.test(p) && !bfWord);
    slots = main3 ? ['lunch','dinner','main3'] : ['breakfast','lunch','dinner'];
  } else slots = ['breakfast','lunch','dinner','main3'];
  const snackSlots = snacks>=2 ? ['amsnack','pmsnack'] : (snacks===1 ? ['amsnack'] : []);
  return {ok:true, slots, snackSlots, kind, n, snacks};
}

/* ---------- dislikes: removables vs dislike meals ---------- */
const stem = w=> w.length>3 ? w.replace(/(IES|ES|S|E)$/,'') : w;
const T_STOP = new Set(['ALL','ANY','KIND','KINDS','FOOD','MEALS','MEAL','THE','IN','ON','FOR','OF','A','WITH','NO','PLEASE','CHOOSE','IS','OK','OKAY','ONLY','EXCEPT','SOME','LESS','MORE','VERY','THICK','FLAVOR']);
const JUNK = /^(LIKES?|PREFERS?|PREFER|ADD|ALWAYS|DO NOT SEND|PLEASE|FOR ALL|ALL MEALS)\b/i;
const GARNISH = new Set(['ONION','GARLIC','MUSHROOM','TOMATO','TOMATOES','CHEESE','CREAM','BROCCOLI','BROCOLI','BROCOCOLI','BROCS','BROC','CAULIFLOWER','COLIFLOWER','CAULI','CARROT','ZUCCHINI','ZUCCHINE','ZUCHINI','PEPPER','CAPSICUM','OLIVE','PICKLE','CORN','CABBAGE','PEAS','BEETROOT','PUMPKIN','NUTS','RAISIN','CORIANDER','JALAPENO','JALPINO','HALOPENO','SPICY','SPICE','SPICES','SAUCE','MAYONAISSE','MAYONNAISE','MAYO','KETCHUP','HUMMUS','LEGUMES','HONEY','GINGER','PEANUT','COCONUT','EGGPLANT','CELERY','AVOCADO','SALT','SOY','BANANA','ASPARAGUS','SPROUTS','CRACKERS','HAZELNUT','LEMON','CARDAMON','VEGGIES','VEGETABLES','TOPPINGS','CUCUMBER','LETTUCE','COLORED','MOHAMMARA'].map(stem));
const BREAD_RE = /\bBREAD\b|SANDWICH|WRAP|BURGER|TOAST|CROISSANT|PITA|\bBUN\b|TORTILLA|SHAWARMA/;
const PASTA_RE = /PASTA|SPAGHETTI|PENNE|FETTUCC|FUSILLI|MACARONI|LASAGN|ALFREDO|NOODLE/;
const GLUTEN_RE = new RegExp(BREAD_RE.source+'|'+PASTA_RE.source+'|COUSCOUS|WAFFLE|PANCAKE|MUFFIN|KUNAFA');
const SEAFOOD_RE = /FISH|SALMON|SHRIMP|PRAWN|TUNA|SEAFOOD|CRAB|CALAMARI|SEA\s*BASS|HAMOUR|MUSSEL|COD\b|TILAPIA/;
const WHITEFISH_RE = /WHITE\s*FISH|SEA\s*BASS|HAMOUR|\bCOD\b|TILAPIA/;
const MEAT_RE = /BEEF|MEAT|STEAK|LAMB|KOFTA|BOLOGNESE|MINCE/;
const DAIRY_RE = /MILK|CHEESE|CREAM|HALLOUMI|YOGURT|YOGHURT|LABNEH|CHEESECAKE|ALFREDO|PUDDING|MAHLABIA|KUNAFA|LAZY CAKE|FETA/;
const dishStems = u=> (u.match(/[A-Z']+/g)||[]).map(stem);

function makeTerm(raw, exc){
  const U = upper(raw);
  const words = (U.match(/[A-Z']+/g)||[]).filter(w=>!T_STOP.has(w) && w.length>1);
  if(!words.length) return null;
  const stems = words.map(stem);
  const W = words.join(' ');
  let rx = null;
  const few = words.length<=2;
  if(few){
    if(/BREAD|SANDWICH|WRAP|TOAST/.test(W)) rx = BREAD_RE;
    else if(/GLUTEN|WHEAT/.test(W)) rx = GLUTEN_RE;
    else if(/PASTA|SPAGHETTI|PENNE|FETTUCCINE|MACARONI|NOODLE/.test(W)) rx = PASTA_RE;
    else if(/WHITE/.test(W) && /FISH/.test(W)) rx = WHITEFISH_RE;
    else if(/SEAFOOD|SHARRY|FISH|MUSSEL|CRAB|CALAMARI/.test(W)) rx = SEAFOOD_RE;
    else if(/SHRIMP|PRAWN/.test(W)) rx = /SHRIMP|PRAWN/;
    else if(/SALMON/.test(W)) rx = /SALMON/;
    else if(/TUNA/.test(W)) rx = /TUNA/;
    else if(/\bMEAT\b|\bMEET\b|BEEF|STEAK|LAMB|KOFTA/.test(W)) rx = MEAT_RE;
    else if(/MILK|DAIRY|LACTOSE|LACTOS/.test(W)) rx = DAIRY_RE;
    else if(words.length===1 && /CHICKEN/.test(W)) rx = /CHICKEN|FAJITA|MADROUBA/;
    else if(words.length===1 && /^EGGS?$/.test(W)) rx = /\bEGG|OMELET|FRITTATA|SHAKSHUKA/;
    else if(words.length===1 && /^(OATMEAL|OATS)$/.test(W)) rx = /OATMEAL|OATS|PORRIDGE/;
    else if(words.length===1 && /^SOUPS?$/.test(W)) rx = /SOUP/;
  }
  const structural = !!rx || !stems.some(s=>GARNISH.has(s));
  const nameHit = u=>{ const ds = dishStems(u); return stems.every(s=>ds.includes(s)); };
  const t = {raw:raw.trim(), words, stems, exc:(exc||[]), structural, rx, nameHit};
  t.test = u=>{
    if(t.exc.some(e=>e && u.includes(e))) return false;
    return rx ? rx.test(u) : nameHit(u);
  };
  // garnish that *is* the dish (e.g. "Mushroom Soup", "Broccoli Soup", "Eggplant Roll") can't be removed
  t.identity = u=> nameHit(u) && (new RegExp('^\\s*(?:KETO\\s+)?'+stems[0]).test(u) || /SOUP|ROLL|PIZZA|CAKE|FRIES|PICKLE|BALLS?|PUDDING/.test(u));
  const lower = t.raw.toLowerCase().replace(/^no\s+/,'');
  t.label = lower;
  return t;
}
function parseTerms(text){
  if(!text) return [];
  const s = String(text).replace(/sweet\s*(?:&|and|n)\s*sour/ig,'sweetsour').replace(/❌|\*|"/g,' ');
  const out = [];
  s.split(/[\n\r,;\/+•]+|\s&\s|\band\b/i).forEach(chunk=>{
    let raw = chunk.trim(); if(!raw) return;
    const exc = [];
    raw = raw.replace(/\(([^)]*)\)/g,(m,inner)=>{
      const a = /(?:ONLY\s+)?([A-Z ]+?)S?\s*(?:IS|ARE)\s*OK(?:AY)?/i.exec(inner) || /EXCEPT\s+([A-Z ]+)/i.exec(inner);
      if(a) exc.push(upper(a[1]).trim().replace(/S$/,''));
      return ' ';
    });
    raw = raw.replace(/\bexcept\s+([A-Za-z ]+)/i,(m,w)=>{ exc.push(upper(w).trim().replace(/S$/,'')); return ' '; });
    raw = raw.replace(/^no\s+/i,'').replace(/\s+/g,' ').trim();
    if(raw.length<3 || JUNK.test(raw)) return;
    const t = makeTerm(raw, exc);
    if(t) out.push(t);
  });
  return out;
}
// allergies + dish dislikes -> substitute the dish; removable ingredients -> keep the dish, write "No X"
function buildCust(c){
  const sub = [], rem = [];
  const route = t=> (t.structural ? sub : rem).push(t);
  parseTerms(c.allergies).forEach(route);
  parseTerms(c.dislikemeals).forEach(t=>sub.push(t));
  parseTerms(c.removables).forEach(route);
  parseTerms(c.dislike).forEach(route); // older single-column files
  return {sub, rem};
}
function nosText(arr){
  const negs = arr.filter(l=>!/^less /.test(l)), less = arr.filter(l=>/^less /.test(l));
  return (negs.length ? 'No '+negs.join(', ') : '') + (less.length ? (negs.length?', ':'')+less.join(', ') : '');
}
// Judge one dish for one customer: substitute it, or serve it with "No X" notes.
function evalDish(dish, cust, slot){
  const u = upper(dish);
  for(const t of cust.sub){ if(t.test(u)) return {sub:true, why:t.label}; }
  const mods = [];
  for(const t of cust.rem){
    const hit = t.nameHit(u);
    if(hit && t.identity(u)) return {sub:true, why:t.label};
    if(hit || slot==='lunch' || slot==='dinner' || slot==='main3'){ if(!mods.includes(t.label)) mods.push(t.label); }
  }
  return {sub:false, mods: mods.slice(0,8)};
}
function pickDish3(main, opt, cust, slot){
  const cands = [{n:main,f:'main'},{n:opt,f:'option'}].filter(c=>c.n);
  if(!cands.length) return {chosen:'', flag:'none', mods:[]};
  const why = [];
  for(const c of cands){
    const e = evalDish(c.n, cust, slot);
    if(!e.sub) return {chosen:c.n, flag:c.f, mods:e.mods, reason: c.f==='option' ? why.join(', ') : ''};
    why.push(e.why);
  }
  return {chosen:'', flag:'manual', mods:[], reason:'every option conflicts ('+why.join(', ')+')'};
}
function pickFiltered(cands, cust, slot, okFn){
  for(const c of cands){
    if(!okFn(c.name)) continue;
    const e = evalDish(c.name, cust, slot);
    if(!e.sub) return {chosen:c.name, flag:c.flag, mods:e.mods, reason:'', src:c.src};
  }
  return null;
}

/* ---------- grammage ---------- */
function defaultGrammage(variant, slot){
  if(variant==='keto') return null;
  if(slot==='lunch') return variant==='hp' ? '150+100' : '100+100';
  if(slot==='dinner') return variant==='hp' ? '150' : '100';
  return null;
}
const R_PC1 = /(\d{2,3})\s*(?:G|GM|GR|GRAMS?)?\s*(?:PRO\w*)\s*(?:[,+&]|AND)?\s*(\d{2,3})\s*(?:G|GM|GR|GRAMS?)?\s*(?:CARB\w*)/;
const R_PC2 = /PRO\w*\s*:?\s*(\d{2,3})\s*(?:G|GM|GR)?\s*(?:[,+&]|AND)?\s*CARB\w*\s*:?\s*(\d{2,3})/;
const R_PC3 = /(\d{2,3})\s*(?:G|GM|GR)?\s*\+\s*(\d{2,3})(?!\d)/;
const R_P1 = /(\d{2,3})\s*(?:G|GM|GR)\s*PRO\w*|PRO\w*\s*:?\s*(\d{2,3})\s*(?:G|GM|GR)/;
function pairIn(s, allowProteinOnly){
  let m = R_PC1.exec(s) || R_PC2.exec(s) || R_PC3.exec(s);
  if(m) return {p:m[1], c:m[2]};
  if(allowProteinOnly){ m = R_P1.exec(s); if(m) return {p:m[1]||m[2]}; }
  return null;
}
// Customer's own grams from free text -> {lunch:{p,c}, dinner:{p,c}}
function parseGrams(text){
  const U = upper(String(text||'').replace(/\s+/g,' '));
  const res = {};
  if(!U) return res;
  const stopRe = /\b(?:DINNE?R?|LUNCH|SNACK|BF|BREAKFAST)\b/;
  const seg = re=>{
    const i = U.search(re); if(i<0) return '';
    const rest = U.slice(i+5); const m = stopRe.exec(rest);
    return U.slice(i, i+5+(m?m.index:rest.length));
  };
  const sL = seg(/\bLUNCH\b/), sD = seg(/\bDINNE/);
  const both = /LUNCH\s*(?:AND|&|\+)\s*DINNER|EVERY|EACH|ALL MEALS|ALL THE MEALS|MAIN MEAL/.test(U);
  const pl = sL ? pairIn(sL, true) : null, pd = sD ? pairIn(sD, true) : null;
  if(pl) res.lunch = pl;
  if(pd) res.dinner = pd;
  if(!pl && !pd || (both && (!pl || !pd))){
    const whole = U.replace(/(BREAKFAST|\bBF\b)[^]*?(?=LUNCH|DINNE|\bM1\b|$)/,' ');
    const w = pairIn(whole, both || (!sL && !sD));
    if(w){ if(!res.lunch) res.lunch = w; if(!res.dinner) res.dinner = w; }
  }
  return res;
}
function gramFor(row, slot, forOrder){
  if(row.variant==='keto') return null;
  const key = slot==='main3' ? 'lunch' : slot;
  const ov = overrides[row.box+'_'+key]; if(ov) return ov;
  if(row.diet==='simple' && !row.custom && !forOrder) return null; // Simple Diet: the sheet shows counts, not grammages
  const def = defaultGrammage(row.variant, key);
  const c = (row.custGram||{})[key];
  if(c && (c.p||c.c)){
    const p = c.p || (def ? def.split('+')[0] : '');
    const carb = c.c || (key==='lunch' && def ? def.split('+')[1] : null);
    return carb ? p+'+'+carb : p;
  }
  return def;
}

/* ---------- per-customer rules (keyed by box number) ----------
   For customers whose own instructions don't fit the generic engine. They
   build the day's assignment directly; anything that can't be satisfied from
   today's menu is flagged for manual handling — never guessed. */
function boxKey(c){ return String(c.newbox||'').replace(/\s+/g,'').toUpperCase(); }
function scheduleBlocks(c){
  const r = CUSTOMER_RULES[boxKey(c)];
  return hasFoodSchedule(c) && !(r && r.foodScheduleMeals);
}
/* menu pools: main + option(s). Macro lunch/dinner carry two options (Classic opt = option 1, HP opt = option 2);
   every customer may be given any of them, so the kitchen can spread the proteins around. */
function poolFor(lineKey, menuSlot){
  const m = menu[lineKey] && menu[lineKey][menuSlot];
  if(!m) return [];
  const names = [m.main];
  if(m.opt && typeof m.opt==='object'){ names.push(m.opt.classic, m.opt.hp); } else names.push(m.opt);
  const seen = new Set(), out = [];
  names.forEach((n,i)=>{
    n = String(n||'').trim();
    if(!n || seen.has(upper(n))) return;
    seen.add(upper(n));
    out.push({name:n, flag: i===0 ? 'main' : 'opt'+i, src:menuSlot});
  });
  return out;
}
// "main3" (third main meal of a no-breakfast plan) is a 2nd lunch or a 2nd dinner
function slotPool(lineKey, slot){
  if(slot==='main3') return poolFor(lineKey,'lunch').concat(poolFor(lineKey,'dinner'));
  return poolFor(lineKey, slot);
}
function slotCandidates(row, lineKey, slot){ return slotPool(lineKey, slot); }

/* variety: spread customers over the available dishes, never repeat a dish or a protein for the same customer
   on the same day when another safe option exists, and never serve a disliked meal */
let varCounts = {};
const PROT = [['chicken',/CHICKEN/],['beef',/BEEF|MEAT|STEAK|BOLOGNESE|LAMB|KOFTA|MINCE/],['salmon',/SALMON/],['shrimp',/SHRIMP|PRAWN/],['tuna',/TUNA/],['fish',/FISH|HAMOUR|SEA\s*BASS|\bCOD\b|TILAPIA/],['turkey',/TURKEY/],['egg',/\bEGG/],['halloumi',/HALLOUMI/]];
function proteinOf(dish){ const u=upper(dish); const h=PROT.find(p=>p[1].test(u)); return h ? h[0] : null; }
function strHash(s){ let h=0; for(let i=0;i<s.length;i++) h=(h*31+s.charCodeAt(i))|0; return Math.abs(h); }
function cmpArr(a,b){ for(let i=0;i<a.length;i++){ if(a[i]!==b[i]) return a[i]-b[i]; } return 0; }
function choosePool(pool, row, slot, lineKey){
  if(!pool.length) return null;
  const ok = [], why = [];
  pool.forEach(c=>{
    const e = evalDish(c.name, row.cust, c.src);
    if(e.sub) why.push(e.why); else ok.push({c, mods:e.mods});
  });
  if(!ok.length) return {chosen:'', flag:'manual', mods:[], reason:'every option conflicts ('+why.join(', ')+')'};
  const t = row.taken || (row.taken = {dishes:{}, prots:{}});
  const lp = slot==='lunch' || slot==='dinner' || slot==='main3';
  let best = null;
  ok.forEach(o=>{
    const U = upper(o.c.name), key = lineKey+'|'+o.c.src+'|'+U, p = proteinOf(o.c.name);
    const score = [t.dishes[U]?1:0, (lp&&p&&t.prots[p])?1:0, varCounts[key]||0, strHash(row.box+row.prodDate+U)%1000];
    if(!best || cmpArr(score,best.score)<0) best = {o, score, key, p, U};
  });
  varCounts[best.key] = (varCounts[best.key]||0)+1;
  t.dishes[best.U] = 1; if(lp && best.p) t.prots[best.p] = 1;
  return {chosen:best.o.c.name, flag:best.o.c.flag, mods:best.o.mods, reason:'', src:best.o.c.src};
}
function menuPick(row, lineKey, slot){
  return choosePool(slotPool(lineKey, slot), row, slot, lineKey);
}
function manualEntry(reason, extra){ return Object.assign({dish:'', flag:'manual', reason, nos:[]}, extra||{}); }
function makeEntry(row, slot, pk){
  if(!pk || (!pk.chosen && pk.flag!=='manual')) return null;
  const isMeal = slot==='lunch'||slot==='dinner'||slot==='main3';
  const src = pk.src || (slot==='main3' ? 'lunch' : slot);
  if(pk.flag==='manual') return manualEntry(pk.reason, isMeal ? {gram:gramFor(row,'lunch')} : {});
  const e = {dish:pk.chosen, flag:pk.flag, reason:pk.reason||'', nos:pk.mods||[], src};
  if(isMeal){ const gk = src==='dinner' ? 'dinner' : 'lunch'; e.gram = gramFor(row, gk); e.ogram = gramFor(row, gk, true); }
  return e;
}
function pickWithRule(cands, okFn){
  for(const c of cands){ if(okFn(c.name)) return {chosen:c.name, flag:c.flag}; }
  return null;
}

const M05_A = '4 Egg White + 60g Halloumi + 3 Small Round Keto Bread';
const M05_B = '2 Egg White + 40g Turkey + 3 Small Round Keto Bread';
const M06_DISLIKES = [
  u=>/TERIYAKI/.test(u) && /CHICKEN/.test(u),
  u=>/SWEET\s*(&|N|AND)?\s*SOUR/.test(u) && /CHICKEN/.test(u),
  u=>/CORIANDER/.test(u) && /FISH/.test(u),
  u=>/TIKKA/.test(u) && !/SALMON|SHRIMP|FISH|BEEF|LAMB/.test(u),
  u=>/TAJ[EI]N|TAGINE/.test(u),
  u=>/BEEF|MEAT|LAMB|STEAK|BOLOGNESE|MINCE|KOFTA/.test(u) && !/BURGER/.test(u),
  u=>PASTA_RE.test(u),
  u=>/PINK\s*RICE/.test(u),
  u=>/OATMEAL|OATS|PORRIDGE/.test(u),
  u=>/HERB/.test(u) && /SALAD/.test(u),
  u=>/CAULIFLOWER|CAULI\b/.test(u),
  u=>/MACHBOO?S|MAJBOO?S|MACHBOUS/.test(u) && !/SHRIMP|SALMON|FISH|BEEF|LAMB/.test(u)
];
const m06Disliked = name=> M06_DISLIKES.some(f=>f(upper(name)));
const M06_MODS = 'salad on side, No cooked veggies';
const M210_PROTEIN = [[/SALMON/,150],[/TUNA/,144],[/CHICKEN/,168],[/SHRIMP|PRAWN/,160],[/STEAK|BEEF|MEAT/,200],[/CRAB/,120]];

const CUSTOMER_RULES = {
  M05: {
    label:'M05', foodScheduleMeals:true,
    build(row, lineKey, weekday){
      const assign = {};
      const A = ['Sunday','Tuesday','Thursday'].includes(weekday), B = ['Saturday','Monday','Wednesday'].includes(weekday);
      const b = A ? M05_A : (B ? M05_B : '');
      if(b) assign.breakfast = {dish:b, flag:'main', reason:'', nos:[]};
      return {assign, note:'Lunch, dinner & snacks are on food schedule — manual', fsAlso:true};
    }
  },
  M06: {
    label:'M06',
    build(row, lineKey, weekday){
      const assign = {};
      const chickenDay = weekday==='Sunday' || weekday==='Thursday';
      const protein = chickenDay ? /CHICKEN/ : /SALMON|SHRIMP|PRAWN/;
      const pName = chickenDay ? 'chicken' : 'salmon/shrimp';
      const gram = {lunch:'150+80', dinner:'150+40'};
      ['lunch','dinner'].forEach(slot=>{
        const p = pickWithRule(slotCandidates(row, lineKey, slot), n=>!m06Disliked(n) && protein.test(upper(n)));
        assign[slot] = p ? {dish:p.chosen, flag:p.flag, reason:'', nos:[], gram:gram[slot], note:M06_MODS}
                         : manualEntry('no allowed '+pName+' meal on today\'s menu', {gram:gram[slot], note:M06_MODS});
      });
      if(row.snacks>0){
        ['amsnack','pmsnack'].forEach(slot=>{
          const p = pickWithRule(slotCandidates(row, lineKey, slot), n=>!m06Disliked(n));
          assign[slot] = p ? {dish:p.chosen, flag:p.flag, reason:'', nos:[]} : manualEntry('every snack option is on the dislike list');
        });
      }
      return {assign};
    }
  },
  M38: {
    label:'M38',
    build(row, lineKey, weekday){
      const assign = {}, extras = [];
      const bk = menuPick(row, lineKey, 'breakfast');
      if(bk){
        assign.breakfast = bk.flag==='manual' ? manualEntry(bk.reason)
          : {dish:bk.chosen, note:'+ Parfait', flag:bk.flag, reason:bk.reason||'', nos:bk.mods||[]};
      }
      const cands = slotCandidates(row, lineKey, 'lunch');
      let okFn = null, need = '';
      if(['Sunday','Tuesday','Thursday'].includes(weekday)){ okFn = n=>PASTA_RE.test(upper(n)); need = 'pasta'; }
      else if(['Wednesday','Saturday'].includes(weekday)){
        okFn = n=>/\bRICE\b|BIRYANI|BUKHARI|MACHBOO?S|MAJBOO?S|KABSA|MANDI/.test(upper(n)) && !/PASTA|POTATO|BREAD|WRAP|\bBUN\b|NOODLE/.test(upper(n));
        need = 'rice-only';
      }
      let L;
      if(okFn){
        const p = pickFiltered(cands, row.cust, 'lunch', okFn);
        L = p ? {dish:p.chosen, flag:p.flag, reason:'', nos:p.mods||[]} : manualEntry('no '+need+' lunch on today\'s menu that fits his dislikes');
      } else {
        const pk = menuPick(row, lineKey, 'lunch');
        L = !pk ? manualEntry('no lunch on menu') : (pk.flag==='manual' ? manualEntry(pk.reason) : {dish:pk.chosen, flag:pk.flag, reason:pk.reason||'', nos:pk.mods||[]});
      }
      L.gram = '170+120'; assign.lunch = L;
      const dk = menuPick(row, lineKey, 'dinner');
      if(dk){
        assign.dinner = dk.flag==='manual' ? manualEntry(dk.reason) : {dish:dk.chosen, flag:dk.flag, reason:dk.reason||'', nos:dk.mods||[]};
        assign.dinner.gram = '170+120';
      }
      extras.push({label:'Snack', dish:'Mediterranean Salad BIG', flag:'main', nos:[]});
      extras.push({label:'Snack', dish:'Halloumi Salad BIG', flag:'main', nos:[]});
      extras.push({label:'Snack', dish:'Oatmeal', flag:'main', nos:[]});
      const sm = menuPick(row, lineKey, 'amsnack');
      if(sm) extras.push(sm.flag==='manual' ? {label:'Sweet snack', dish:'', flag:'manual', reason:sm.reason, nos:[]} : {label:'Sweet snack', dish:sm.chosen, flag:sm.flag, nos:sm.mods||[], src:sm.src});
      extras.push({label:'Snack', dish:'Berries — 50g blueberry + 50g strawberry', flag:'main', nos:[]});
      return {assign, extras};
    }
  },
  // SD047: 3 main meals at 150+120 from the Simple Diet menu + 1 salad taken from the Macro menu at 150. No dislike text beside her box.
  SD047: {
    label:'SD047',
    build(row, lineKey){
      const assign = {}, extras = [];
      ['lunch','dinner','main3'].forEach(slot=>{
        const pk = menuPick(row, lineKey, slot);
        if(!pk) return;
        if(pk.flag==='manual'){ assign[slot] = manualEntry(pk.reason, {gram:'150+120'}); return; }
        assign[slot] = {dish:pk.chosen, flag:pk.flag, reason:'', nos:[], gram:'150+120', src:pk.src};
      });
      const cands = [];
      ['lunch','dinner','amsnack','pmsnack'].forEach(s=> poolFor('macro', s).forEach(c=>{ if(/SALAD/.test(upper(c.name))) cands.push(c); }));
      const pk = choosePool(cands, row, 'amsnack', 'macro');
      extras.push(pk && pk.flag!=='manual'
        ? {label:'Salad', dish:pk.chosen, flag:pk.flag, nos:[], gram:'150', line:'macro', src:pk.src}
        : {label:'Salad', dish:'', flag:'manual', reason:'no salad on today\'s Macro menu', nos:[], line:'macro'});
      return {assign, extras};
    }
  },
  // 3 main meals (no breakfast) + 1 salad; carbs 165g, protein grams depend on the protein
  M210: {
    label:'M210',
    build(row, lineKey){
      const assign = {}, extras = [];
      ['lunch','dinner','main3'].forEach(slot=>{
        const pk = menuPick(row, lineKey, slot);
        if(!pk) return;
        if(pk.flag==='manual'){ assign[slot] = manualEntry(pk.reason); return; }
        const hit = M210_PROTEIN.find(([re])=>re.test(upper(pk.chosen)));
        assign[slot] = {dish:pk.chosen, flag:pk.flag, src:pk.src, reason:pk.reason||'', nos:pk.mods||[], gram:(hit?hit[1]:150)+'+165', note: hit ? '' : 'protein weight not set for this dish — check'};
      });
      const salads = [];
      ['pmsnack','amsnack'].forEach(s=> slotCandidates(row, lineKey, s).forEach(c=>{ if(/SALAD/.test(upper(c.name))) salads.push(c); }));
      const sp = pickFiltered(salads, row.cust, 'amsnack', ()=>true);
      extras.push(sp ? {label:'Salad', dish:sp.chosen, flag:sp.flag, nos:sp.mods||[], src:sp.src} : {label:'Salad', dish:'', flag:'manual', reason:'no salad on today\'s snack menu', nos:[]});
      return {assign, extras};
    }
  },
  // only fish & chicken dishes, carbs alternate potato / rice; breakfast = eggs; 3 meals, no snacks
  M217: {
    label:'M217',
    build(row, lineKey){
      const assign = {};
      const par = Math.floor(Date.parse(row.prodDate+'T00:00:00Z')/86400000) % 2;
      const bk = pickFiltered(slotCandidates(row, lineKey, 'breakfast'), row.cust, 'breakfast', n=>/\bEGG|OMELET/.test(upper(n)));
      assign.breakfast = bk ? {dish:bk.chosen, flag:bk.flag, reason:'', nos:bk.mods||[]} : manualEntry('no egg breakfast on today\'s menu');
      const okP = n=>/CHICKEN|FISH|SALMON|SEA\s*BASS|HAMOUR|\bCOD\b|TILAPIA|TUNA/.test(upper(n));
      ['lunch','dinner'].forEach((slot,i)=>{
        const carb = ((par + i) % 2) ? 'Rice' : 'Potato';
        const p = pickFiltered(slotCandidates(row, lineKey, slot), row.cust, slot, okP);
        assign[slot] = p ? {dish:p.chosen, flag:p.flag, reason:'', nos:p.mods||[], gram:'150+100', carb, note:'carb: '+carb}
                         : manualEntry('no chicken/fish dish on today\'s menu', {gram:'150+100', carb});
      });
      return {assign};
    }
  }
};

/* ---------- carbs summary ---------- */
const CARB_DEFS = [
  ['Pink Rice', /PINK\s*RICE/, true],
  ['Brown Rice', /BROWN\s*RICE/, true],
  ['Sweet Potato', /SWEET\s*POTATO/, true],
  ['Rice', /\bRICE\b|BUKHARI|BIRYANI|MACHBOO?S|MAJBOO?S|MANDI|KABSA|MADBI/, false],
  ['Pasta', /PASTA|SPAGHETTI|PENNE|FETTUCC|FUSILLI|MACARONI|LASAGN|NOODLE|ALFREDO/, false],
  ['Potato', /POTATO|WEDGES|FRIES|MASH/, false],
  ['Bread / Bun / Wrap / Waffle', /BREAD|\bBUN\b|WRAP|TORTILLA|CROISSANT|TOAST|PITA|SANDWICH|WAFFLE|PANCAKE/, false],
  ['Quinoa', /QUINOA/, false],
  ['Couscous', /COUSCOUS/, false],
  ['Oats / Oatmeal', /OATMEAL|OATS|GRANOLA/, false],
  ['Barley', /BARLEY/, false]
];
function carbsIn(dish){
  const u = upper(dish);
  const u2 = u.replace(/PINK\s*RICE|BROWN\s*RICE|SWEET\s*POTATO/g,' ');
  const out = [];
  CARB_DEFS.forEach(([name, re, specific])=>{ if(re.test(specific ? u : u2)) out.push(name); });
  return out;
}
function allEntries(r){ return SLOT_ORDER.map(s=>r.assign[s]).concat(r.extras||[]); }
function collectCarbs(rows){
  const carbs = {};
  function add(name, dish, grams){
    if(!carbs[name]) carbs[name] = {portions:0, grams:0, dishes:{}};
    carbs[name].portions++; carbs[name].grams += grams; carbs[name].dishes[dish] = (carbs[name].dishes[dish]||0)+1;
  }
  rows.forEach(r=>{
    allEntries(r).forEach(a=>{
      if(!a || !a.dish || a.flag==='manual') return;
      const found = a.carb ? [a.carb] : carbsIn(a.dish);
      if(!found.length) return;
      const gm = /^(\d+)\+(\d+)/.exec(a.gram||'');
      const g = gm ? (+gm[2])/found.length : 0;
      found.forEach(n=>add(n, a.dish, g));
    });
  });
  return carbs;
}

/* ---------- run ---------- */
function runEngine(){
  const prodDate = prodDateInput.value;
  const weekday = weekdayFor(prodDate);
  const results = [];
  varCounts = {};

  allCustomers().forEach(c=>{
    const row = {box:c.newbox||'', name:c.name||'', list:c.list||'', raw:c, prodDate};
    const diet = normDiet(c.list);
    row.diet = diet;
    const st = statusCheck(c);
    if(!st.ok){ row.provide=false; row.excludeReason = st.reason; results.push(row); return; }
    const dOk = dateOk(c, prodDate);
    if(!dOk.ok){ row.provide=false; row.excludeReason = dOk.reason; results.push(row); return; }
    if(frozenOn(c, prodDate)){ row.provide=false; row.excludeReason = 'Frozen on this date'; results.push(row); return; }
    if(diet==='simple' && weekday==='Friday'){ row.provide=false; row.excludeReason = 'Simple Diet is closed on Fridays'; results.push(row); return; }
    if(offDayHit(c, weekday)){ row.provide=false; row.excludeReason = 'Off day ('+weekday+')'; results.push(row); return; }
    // food schedule: they are due today but selected by hand — listed at the end of the sheet
    if(scheduleBlocks(c)){ row.provide=false; row.excludeReason = 'Food schedule — manual'; row.skipManual = true; row.fs = true; results.push(row); return; }

    const planText = c.mealplan || '';
    const noteText = [c.custnotes, c.packnote, c.note].filter(Boolean).join(' | ');
    row.variant = diet==='keto' ? 'keto' : (classifyVariant(planText)==='hp' ? 'hp' : 'classic');
    row.cust = buildCust(c);
    row.custGram = parseGrams(noteText);
    row.custom = !!(row.custGram.lunch || row.custGram.dinner);
    row.custNoteText = noteText;
    row.snacks = snackCount(planText);

    const lineKey = diet==='app' ? null : diet;
    row.assign = {}; row.extras = [];
    if(!lineKey || !menu[lineKey]){ row.provide=true; row.appOnly=true; results.push(row); return; }

    const rule = CUSTOMER_RULES[boxKey(c)];
    if(rule){
      row.provide = true; row.rule = rule.label;
      const ra = rule.build(row, lineKey, weekday);
      row.assign = ra.assign; row.extras = ra.extras || []; row.ruleNote = ra.note || ''; row.fsAlso = !!ra.fsAlso;
      results.push(row); return;
    }

    const comp = composition(planText, noteText+' '+(c.removables||''));
    if(!comp.ok){ row.provide=false; row.excludeReason = comp.reason+' — manual'; row.skipManual = true; results.push(row); return; }
    row.provide = true; row.kind = comp.kind;
    comp.slots.concat(comp.snackSlots).forEach(slot=>{
      let e = makeEntry(row, slot, menuPick(row, lineKey, slot));
      // one-snack packages: AM snack, or the PM snack when the AM one is blocked
      if(slot==='amsnack' && comp.snacks===1 && (!e || e.flag==='manual')){
        const e2 = makeEntry(row, 'pmsnack', menuPick(row, lineKey, 'pmsnack'));
        if(e2 && e2.flag!=='manual'){ row.assign.pmsnack = e2; return; }
      }
      if(e) row.assign[slot] = e;
    });
    results.push(row);
  });

  runResults = {prodDate, weekday, results};
  renderRunSummary();
  renderDetail();
  renderSheet();
  refreshEligCount();
}
$('#runBtn').addEventListener('click', runEngine);

function refreshEligCount(){ refreshEligCount0(); if(typeof renderHome==='function') renderHome(); }
function refreshEligCount0(){
  if(!allCustomers().length){ $('#eligCount').textContent = '0'; return; }
  const prodDate = prodDateInput.value;
  const weekday = weekdayFor(prodDate);
  let n=0;
  allCustomers().forEach(c=>{
    if(!statusCheck(c).ok) return;
    if(frozenOn(c, prodDate)) return;
    if(scheduleBlocks(c)) return;
    if(!dateOk(c, prodDate).ok) return;
    if(normDiet(c.list)==='simple' && weekday==='Friday') return;
    if(offDayHit(c, weekday)) return;
    n++;
  });
  $('#eligCount').textContent = String(n);
}
refreshEligCount();

function renderRunSummary(){
  const el = $('#runSummary');
  if(!runResults){ el.innerHTML = '<div class="empty">Not run yet.</div>'; return; }
  const r = runResults.results;
  const provide = r.filter(x=>x.provide);
  const excluded = r.filter(x=>!x.provide);
  const manual = provide.filter(x=>allEntries(x).some(a=>a && a.flag==='manual'));
  const fs = r.filter(x=>x.fs || x.fsAlso);
  const unclear = excluded.filter(x=>x.skipManual && !x.fs);
  el.innerHTML = `
    <div class="sum-grid">
      <div class="sum-tile"><div class="num">${r.length}</div><div class="lbl">Total customers</div></div>
      <div class="sum-tile"><div class="num" style="color:var(--ok)">${provide.length}</div><div class="lbl">Provide today</div></div>
      <div class="sum-tile"><div class="num" style="color:var(--ink-dim)">${excluded.length}</div><div class="lbl">Excluded</div></div>
      <div class="sum-tile"><div class="num" style="color:var(--bad)">${manual.length}</div><div class="lbl">Need manual pick</div></div>
      <div class="sum-tile"><div class="num" style="color:var(--warn)">${fs.length}</div><div class="lbl">Food schedule (manual)</div></div>
      <div class="sum-tile"><div class="num" style="color:var(--warn)">${unclear.length}</div><div class="lbl">Plan unclear (manual)</div></div>
    </div>
    <p class="foot-note" style="margin-top:12px;">${runResults.weekday}, ${runResults.prodDate}. Full breakdown is in the <strong>Customer list</strong> and <strong>Production sheet</strong> tabs.</p>
  `;
}

function escHtml(s){ return String(s==null?'':s).replace(/[&<>"]/g, c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;'}[c])); }

function renderDetail(){
  const host = $('#detailTable');
  if(!runResults){ host.innerHTML = '<div class="empty">Run the engine first.</div>'; return; }
  const rows = runResults.results;
  let html = `<div class="tbl-wrap"><table><thead><tr>
    <th>Box</th><th>Name</th><th>Line</th><th>Variant</th><th>Status</th>
    ${SLOT_ORDER.map(s=>`<th>${SLOT_LABEL[s]}</th>`).join('')}
    <th>Notes</th>
  </tr></thead><tbody>`;
  rows.forEach(row=>{
    if(!row.provide){
      html += `<tr><td class="mono">${escHtml(row.box)}</td><td>${escHtml(row.name)}</td><td>${escHtml(row.list)}</td><td>—</td>
        <td colspan="8"><span class="status-pill ${row.skipManual?'pill-warn':'pill-dim'}">${escHtml(row.excludeReason)}</span></td></tr>`;
      return;
    }
    const cells = SLOT_ORDER.map(slot=>{
      const a = row.assign[slot];
      if(!a) return '<td class="foot-note">—</td>';
      const gramTxt = a.gram ? ` <span class="badge">${escHtml(a.gram)}</span>` : '';
      let txt;
      if(a.flag==='manual') txt = `<span class="status-pill pill-bad">manual pick</span><div class="foot-note">${escHtml(a.reason||'')}</div>`;
      else {
        txt = `${/^opt/.test(a.flag||'') ? '<span class="badge">'+escHtml(a.flag.replace('opt','option '))+'</span> ' : ''}${escHtml(a.dish||'—')}${gramTxt}${slot==='main3'&&a.src?` <span class="foot-note">(extra ${escHtml(a.src)})</span>`:''}`;
      }
      if(a.note) txt += `<div class="foot-note">${escHtml(a.note)}</div>`;
      if(a.nos && a.nos.length) txt += `<div class="foot-note" style="color:var(--warn)">${escHtml(nosText(a.nos))}</div>`;
      return `<td>${txt}</td>`;
    }).join('');
    let notes = row.custom ? `<span class="status-pill pill-warn" title="${escHtml(row.custNoteText)}">own grams</span>` : '';
    if(row.rule) notes += `<span class="status-pill pill-ok">rule ${escHtml(row.rule)}</span>`;
    if(row.fsAlso) notes += `<span class="status-pill pill-warn">food schedule</span>`;
    if(row.ruleNote) notes += `<div class="foot-note">${escHtml(row.ruleNote)}</div>`;
    if(row.extras && row.extras.length) notes += `<div class="foot-note">+ ${row.extras.map(x=>escHtml(x.dish||'manual')).join('; ')}</div>`;
    html += `<tr>
      <td class="mono">${escHtml(row.box)}</td><td>${escHtml(row.name)}</td><td>${escHtml(row.list)}</td>
      <td>${escHtml(row.variant||'')}</td><td><span class="status-pill pill-ok">provide</span></td>
      ${cells}
      <td>${notes}</td>
    </tr>`;
  });
  html += '</tbody></table></div>';
  host.innerHTML = html;
}

/* ---------- production sheet (layout follows the AllSheetMeals template) ---------- */
const SHEET_LINES = [
  {key:'keto', label:'KETO', slots:[
    {slot:'breakfast', rows:['BREAKFAST','OPTION BREAKFAST']},
    {slot:'lunch',     rows:['LUNCH','OPTION LUNCH']},
    {slot:'dinner',    rows:['DINNER','OPTION DINNER']},
    {slot:'amsnack',   rows:['SNACK 1','OPTION SNACK 1']},
    {slot:'pmsnack',   rows:['SNACK 2','OPTION SNACK 2']}]},
  {key:'macro', label:'MACRO', slots:[
    {slot:'breakfast', rows:['BREAKFAST','BREAKFAST OPTION'], other:'BREAKFAST OTHER OPTIONS'},
    {slot:'lunch',     rows:['LUNCH','LUNCH OPTION 1','LUNCH OPTION 2'], other:'LUNCH OTHER OPTION'},
    {slot:'dinner',    rows:['DINNER','DINNER OPTION 1','DINNER OPTION 2'], other:'DINNER OTHER OPTION'},
    {slot:'amsnack',   rows:['AM SNACK','AM SNACK OPTION'], other:'AM OTHER SNACK OPTION'},
    {slot:'pmsnack',   rows:['PM SNACK','PM SNACK OPTION'], other:'PM OTHER SNACK OPTION'}]},
  {key:'simple', label:'SIMPLE DIET', slots:[
    {slot:'breakfast', rows:['BREAKFAST','OPTION BREAKFAST']},
    {slot:'lunch',     rows:['LUNCH','OPTION LUNCH']},
    {slot:'dinner',    rows:['DINNER','OPTION DINNER']},
    {slot:'amsnack',   rows:['AM SNACK','OPTION AM SNACK']},
    {slot:'pmsnack',   rows:['PM SNACK','OPTION PM SNACK']}]}
];
function entryTarget(r, slot, a){
  return {line: a.line || r.diet, ms: a.src || (slot==='main3' ? 'lunch' : slot)};
}
function buildSheet(results){
  const lines = SHEET_LINES.map(L=>({
    key:L.key, label:L.label, total:0, rem:0, special:{},
    slots: L.slots.map(s=>{
      const pool = poolFor(L.key, s.slot);
      return {slot:s.slot, otherLabel:s.other||'OTHER / SPECIAL', hasOtherRow:!!s.other, others:{},
        rows: s.rows.map((lab,i)=>{ const p = pool.find(c=>c.flag===(i===0?'main':'opt'+i)); return {label:lab, dish:p?p.name:'', items:[]}; })};
    })
  }));
  const remaining = [];
  const remMap = {};
  const addRem = (r, why)=>{
    if(!remMap[r.box]){ remMap[r.box] = {box:r.box, diet:r.diet, name:r.name, raw:r.raw, reasons:[]}; remaining.push(remMap[r.box]); }
    if(!remMap[r.box].reasons.includes(why)) remMap[r.box].reasons.push(why);
  };
  results.forEach(r=>{
    if(r.fs){ addRem(r, 'Food schedule'); return; }
    if(r.fsAlso) addRem(r, 'Food schedule (lunch, dinner, snacks)');
    if(!r.provide){ if(r.skipManual) addRem(r, r.excludeReason.replace(/ — manual$/,'')); return; }
    const put = (slot, a)=>{
      if(!a) return;
      if(a.flag==='manual'){ addRem(r, (SLOT_LABEL[slot]||a.label||slot)+': '+(a.reason||'no safe option')); return; }
      const t = entryTarget(r, slot, a);
      const L = lines.find(x=>x.key===t.line); if(!L) return;
      const S = L.slots.find(x=>x.slot===t.ms);
      const item = {box:r.box, nos:a.nos||[], gram:a.gram||'', note:a.note||'', dish:a.dish, og:a.ogram||a.gram||''};
      if(!S){ const k = upper(a.dish); (L.special[k] = L.special[k] || {dish:a.dish, items:[]}).items.push(item); L.total++; if(item.nos.length) L.rem++; return; }
      const pool = poolFor(t.line, t.ms);
      const p = pool.find(c=>upper(c.name)===upper(a.dish));
      const idx = p ? (p.flag==='main' ? 0 : +p.flag.slice(3)) : -1;
      if(idx>=0 && idx<S.rows.length) S.rows[idx].items.push(item);
      else { const k = upper(a.dish); (S.others[k] = S.others[k] || {label:S.otherLabel, dish:a.dish, items:[]}).items.push(item); }
      L.total++; if(item.nos.length) L.rem++;
    };
    SLOT_ORDER.forEach(s=> put(s, r.assign[s]));
    (r.extras||[]).forEach(x=> put('extra', x));
  });
  const order = ['keto','macro','simple'];
  remaining.sort((a,b)=> (order.indexOf(a.diet)<0?9:order.indexOf(a.diet)) - (order.indexOf(b.diet)<0?9:order.indexOf(b.diet)));
  return {lines, remaining};
}
// boxes of one row, grouped by identical grammage (+ note); customers with removable dislikes go last in their group
function itemGroups(items){
  const g = {};
  items.forEach(it=>{
    const k = it.gram+'|'+it.note;
    (g[k] = g[k] || {gram:it.gram, note:it.note, clean:[], dis:[]})[it.nos.length ? 'dis' : 'clean'].push(it);
  });
  return Object.values(g).sort((a,b)=> a.gram.localeCompare(b.gram,undefined,{numeric:true}) || a.note.localeCompare(b.note));
}
const grpPrefix = g=> (g.gram||g.note) ? [g.gram, g.note].filter(Boolean).join(' ')+'—' : '';
function groupsText(items){
  return itemGroups(items).map(g=> grpPrefix(g).replace(/—$/,'--') + g.clean.map(i=>i.box).concat(g.dis.map(i=>i.box+' ('+nosText(i.nos)+')')).join(' / ')).join('  |  ');
}
function groupsHtml(items){
  return itemGroups(items).map(g=>{
    const clean = g.clean.map(i=>escHtml(i.box)).join(' / ');
    const dis = g.dis.map(i=>`<span class="status-pill pill-warn">${escHtml(i.box)}, ${escHtml(nosText(i.nos))}</span>`).join(' ');
    return `<div>${grpPrefix(g)?`<span class="foot-note">${escHtml(grpPrefix(g))}</span>`:''}${clean}${clean&&dis?' &nbsp;':''}${dis}</div>`;
  }).join('');
}
function sheetRows(L){
  const out = [];
  L.slots.forEach(S=>{
    S.rows.forEach(r=> out.push({label:r.label, dish:r.dish, items:r.items}));
    const ks = Object.keys(S.others);
    if(ks.length) ks.forEach(k=> out.push({label:S.otherLabel, dish:S.others[k].dish, items:S.others[k].items, other:true}));
    else if(S.hasOtherRow) out.push({label:S.otherLabel, dish:'', items:[]});
  });
  Object.keys(L.special).forEach(k=> out.push({label:'SPECIAL ADD-ON', dish:L.special[k].dish, items:L.special[k].items, other:true}));
  return out;
}

function renderSheet(){
  const host = $('#sheetTable'), totalsHost = $('#sheetTotals');
  if(!runResults){ host.innerHTML = '<div class="empty">Run the engine first.</div>'; totalsHost.innerHTML=''; return; }
  const rows = runResults.results.filter(r=>r.provide);
  const sheet = buildSheet(runResults.results);
  runResults.sheet = sheet;
  const grand = sheet.lines.reduce((a,L)=>a+L.total,0), grandRem = sheet.lines.reduce((a,L)=>a+L.rem,0);
  totalsHost.innerHTML = `<div class="sum-grid" style="margin-bottom:14px;">
    ${sheet.lines.map(L=>`<div class="sum-tile"><div class="num">${L.total}</div><div class="lbl">${L.label} meals</div></div>`).join('')}
    <div class="sum-tile"><div class="num" style="color:var(--accent-text)">${grand}</div><div class="lbl">Total meals</div></div>
    <div class="sum-tile"><div class="num" style="color:var(--warn)">${grandRem}</div><div class="lbl">With removable dislikes</div></div>
    <div class="sum-tile"><div class="num" style="color:var(--bad)">${sheet.remaining.length}</div><div class="lbl">Boxes to do by hand</div></div>
  </div>`;
  let html = `<p class="foot-note" style="margin-bottom:10px;">${escHtml(runResults.weekday)} ${escHtml(runResults.prodDate)}. Customers with removable dislikes are listed last in each row, with the dislike beside the box.</p>`;
  sheet.lines.forEach(L=>{
    const closed = L.key==='simple' && runResults.weekday==='Friday';
    html += `<h3 style="margin:18px 0 6px;">${L.label} <span class="foot-note">· ${L.total} meals${L.rem?`, ${L.rem} with removable dislikes`:''}${closed?' · closed on Fridays':''}</span></h3>
      <div class="tbl-wrap"><table><thead><tr><th style="width:150px;">Row</th><th>Meal</th><th style="width:48px;">Qty</th><th style="width:64px;">Dislikes</th><th>Boxes${L.key==='keto'?'':(L.key==='macro'?' (grammage—box)':'')}</th></tr></thead><tbody>`;
    sheetRows(L).forEach(r=>{
      html += `<tr><td><strong>${escHtml(r.label)}</strong></td><td>${r.dish?escHtml(r.dish):'<span class="foot-note">—</span>'}</td>
        <td class="mono">${r.items.length||''}</td><td class="mono">${r.items.filter(i=>i.nos.length).length||''}</td><td class="mono">${groupsHtml(r.items)}</td></tr>`;
    });
    html += '</tbody></table></div>';
  });

  // everyone who is not on the lists above: food schedule, plan unclear, or no meal possible because of dislikes
  const rem = sheet.remaining;
  html += `<h3 style="margin:22px 0 6px;">Boxes to select by hand <span class="foot-note">· ${rem.length}</span></h3>
    <div class="tbl-wrap"><table><tbody><tr><td class="mono" style="white-space:normal;">${rem.length ? rem.map(x=>escHtml(x.box)).join(' / ') : '—'}</td></tr></tbody></table></div>`;
  if(rem.length){
    html += `<details style="margin-top:8px;"><summary>Why each box is here</summary><div class="tbl-wrap"><table><thead><tr><th>Box</th><th>Line</th><th>Reason</th></tr></thead><tbody>${
      rem.map(x=>`<tr><td class="mono">${escHtml(x.box)}</td><td>${escHtml(LINE_LABEL[x.diet]||'')}</td><td class="foot-note">${escHtml(x.reasons.join('; '))}</td></tr>`).join('')}</tbody></table></div></details>`;
  }
  host.innerHTML = html;
  renderCarbs(rows);
  renderCounts(sheet); renderManual(); renderOrders();
}

function fmtGrams(g){ return g>=1000 ? (g/1000).toFixed(2).replace(/\.?0+$/,'')+' kg' : (g? Math.round(g)+' g' : '—'); }
function renderCarbs(rows){
  const host = $('#carbTable');
  const carbs = collectCarbs(rows);
  const names = Object.keys(carbs).sort((a,b)=>carbs[b].portions-carbs[a].portions);
  runResults.carbs = carbs;
  if(!names.length){ host.innerHTML = '<div class="empty">No carbs detected in today\'s dishes.</div>'; return; }
  let html = `<div class="tbl-wrap"><table><thead><tr><th>Carb</th><th>Portions</th><th>Carb weight (L/D)</th><th>Used in</th></tr></thead><tbody>`;
  names.forEach(n=>{
    const c = carbs[n];
    const used = Object.entries(c.dishes).sort((a,b)=>b[1]-a[1]).map(([d,q])=>`${escHtml(d)} ×${q}`).join('; ');
    html += `<tr><td><strong>${escHtml(n)}</strong></td><td class="mono">${c.portions}</td><td class="mono">${fmtGrams(c.grams)}</td><td class="foot-note">${used}</td></tr>`;
  });
  html += '</tbody></table></div><p class="foot-note" style="margin-top:8px;">Carb weight = the second number of each lunch/dinner grammage (e.g. 150+<strong>100</strong>). Carbs are detected from dish names; breakfast/snack carbs are counted as portions only.</p>';
  host.innerHTML = html;
}

/* ---------- CSV export (clipboard, since downloads are blocked in this preview) ---------- */
function toCSV(rows){
  return rows.map(r=>r.map(c=>{
    const s = String(c==null?'':c);
    return /[",\n]/.test(s) ? '"'+s.replace(/"/g,'""')+'"' : s;
  }).join(',')).join('\n');
}
async function copyCSV(text, btn){
  try{ await navigator.clipboard.writeText(text); btn.textContent='Copied ✓'; }
  catch(e){
    const ta=document.createElement('textarea'); ta.value=text; document.body.appendChild(ta); ta.select();
    try{ document.execCommand('copy'); btn.textContent='Copied ✓'; }catch(e2){ btn.textContent='Select & copy manually'; }
    ta.remove();
  }
  setTimeout(()=>{ btn.textContent = btn.dataset.label; }, 1800);
}
/* ---------- delivering a file ----------
   Inside Claude the viewer is asked to save the file (downloads capability). In an installed app the browser downloads it.
   Anywhere else the text version is copied to the clipboard. */
function flash(btn, txt){ btn.textContent = txt; setTimeout(()=>{ btn.textContent = btn.dataset.label || btn.textContent; }, 1800); }
async function deliver(filename, data, btn, csvFallback){
  const csv = typeof data==='string' ? data : (csvFallback||'');
  if(typeof data==='string') data = '﻿' + data; // BOM so Excel reads the file as UTF-8
  try{
    const dl = (window.claude && window.claude.use) ? await window.claude.use('downloads') : null;
    if(dl){ await dl.save({filename, data}); flash(btn,'Saved ✓'); return; }
  }catch(err){
    if(err && err.code==='declined') return;
  }
  try{
    const url = URL.createObjectURL(data instanceof Blob ? data : new Blob([data]));
    const a = document.createElement('a'); a.href = url; a.download = filename; document.body.appendChild(a); a.click(); a.remove();
    setTimeout(()=>URL.revokeObjectURL(url), 3000); flash(btn,'Downloaded ✓'); return;
  }catch(err){}
  copyCSV(csv, btn);
}

$('#exportDetail').dataset.label = 'Export CSV';
$('#exportSheet').dataset.label = 'Export CSV';
$('#exportDetail').addEventListener('click', (e)=>{
  if(!runResults) return;
  const header = ['Box','Name','Line','Variant','Status', ...SLOT_ORDER.map(s=>SLOT_LABEL[s]), 'Grammage(L/D)','Notes'];
  const data = [header];
  runResults.results.forEach(r=>{
    if(!r.provide){ data.push([r.box,r.name,r.list,'',r.excludeReason, '','','','','','','']); return; }
    const cells = SLOT_ORDER.map(s=>{
      const a = r.assign[s]; if(!a) return '';
      if(a.flag==='manual') return 'MANUAL: '+(a.reason||'');
      return a.dish + (a.nos&&a.nos.length ? ' ['+nosText(a.nos)+']' : '') + (a.note?' ('+a.note+')':'');
    });
    const grams = [r.assign.lunch?.gram||'', r.assign.dinner?.gram||''].filter(Boolean).join(' / ');
    const notes = [r.custom?('OWN GRAMS: '+r.custNoteText):'', r.rule?('RULE '+r.rule):'', r.ruleNote||'', (r.extras&&r.extras.length)?('ADD-ONS: '+r.extras.map(x=>x.dish||'manual').join('; ')):''].filter(Boolean).join(' | ');
    data.push([r.box, r.name, r.list, r.variant, 'provide', ...cells, grams, notes]);
  });
  copyCSV(toCSV(data), e.target);
});
$('#exportSheet').addEventListener('click', (e)=>{
  if(!runResults || !runResults.sheet) return;
  const sh = runResults.sheet;
  const data = [['Line','Row','Meal','Qty','Boxes (grammage—box; dislikes last)','With dislikes']];
  sh.lines.forEach(L=> sheetRows(L).forEach(r=> data.push([L.label, r.label, r.dish, r.items.length||'', groupsText(r.items), r.items.filter(i=>i.nos.length).length||''])));
  data.push([]);
  data.push(['BOXES TO SELECT BY HAND', sh.remaining.length, sh.remaining.map(x=>x.box).join(' / ')]);
  sh.remaining.forEach(x=> data.push(['', x.box, x.name, '', x.reasons.join('; ')]));
  if(runResults.carbs){
    data.push([]); data.push(['CARBS TO PREPARE','Portions','Carb weight (L/D, g)','Used in']);
    Object.entries(runResults.carbs).sort((a,b)=>b[1].portions-a[1].portions).forEach(([n,c])=>{
      data.push([n, c.portions, Math.round(c.grams), Object.entries(c.dishes).map(([d,q])=>d+' x'+q).join('; ')]);
    });
  }
  deliver('production-sheet-'+runResults.prodDate+'.csv', toCSV(data), e.target);
});


/* ======================= customer data sheet (grid) ======================= */
const HK_ORDER = {keto:0, macro:1, simple:2, app:3};
function sortedRoster(){
  return roster.slice().sort((a,b)=> (HK_ORDER[a.hk]-HK_ORDER[b.hk]) || a.box.localeCompare(b.box,undefined,{numeric:true}));
}
const DAY_LETTER = {Sunday:'S',Monday:'M',Tuesday:'T',Wednesday:'W',Thursday:'T',Friday:'F',Saturday:'S'};
function locText(l){
  l = l || {};
  const a = [l.zone,l.street,l.building].filter(Boolean).join(', ');
  const d = (l.fromDay||l.toDay) ? ' · '+(l.fromDay||'').slice(0,3)+'–'+(l.toDay||'').slice(0,3) : '';
  const t = (l.fromDate||l.toDate) ? ' · '+(l.fromDate||'…')+' → '+(l.toDate||'…') : '';
  return a ? a+d+t : '';
}
function rosterTotals(today){
  const keys = ['keto','macro','simple','app'];
  const mk = ()=>({keto:0,macro:0,simple:0,app:0,total:0});
  const all = mk(), act = mk(), frz = mk();
  roster.forEach(e=>{
    const st = statusInfo(e,today)[0], k = keys.includes(e.hk) ? e.hk : 'app';
    all[k]++; all.total++;
    if(st==='Active'){ act[k]++; act.total++; }
    if(st==='Frozen'){ frz[k]++; frz.total++; }
  });
  return {all, act, frz};
}
let pendingDelete = null;
function renderGrid(){
  const host = $('#gridHost'); if(!host) return;
  const today = todayISO();
  if(!roster.length){ host.innerHTML = '<div class="empty">The customer data sheet is empty. Choose <strong>+ Register customer</strong>, or load a workbook above.</div>'; return; }
  const q = str($('#gSearch').value).toLowerCase(), fh = $('#gHk').value, fs = $('#gStatus').value;
  const view = sortedRoster().filter(e=>{
    if(fh!=='all' && e.hk!==fh) return false;
    if(fs!=='all' && !statusInfo(e,today)[0].startsWith(fs)) return false;
    return !q || (e.name+' '+e.box+' '+e.contact).toLowerCase().includes(q);
  });
  const cols = ['A','B','C','D','E','F','G','H','I','J','K','L','M','N','','O','P','R','R',''];
  const names = ['HK','Status','Contact','Box','Customer name','Meal plan package','Allergies','Dislikes (ingredients)','Dislike meals','Customer notes','Pack note','Duration','Joining date','Due date','Freeze','Off days','Remaining','Location A','Location B',''];
  let html = `<div class="tbl-wrap xgrid"><table><thead>
    <tr class="letters">${cols.map(c=>`<th>${c}</th>`).join('')}</tr>
    <tr class="names">${names.map(n=>`<th>${n}</th>`).join('')}</tr></thead><tbody>`;
  view.forEach(e=>{
    const st = statusInfo(e,today), due = dueEff(e,today), rem = remainingDays(e,today);
    html += `<tr>
      <td>${escHtml(PKG_LABEL[e.hk]||'')}</td>
      <td><span class="status-pill ${st[1]}">${escHtml(st[0])}</span></td>
      <td class="mono">${escHtml(e.contact)}</td>
      <td class="mono"><strong>${escHtml(e.box)}</strong></td>
      <td class="wrapc">${escHtml(e.name)}</td>
      <td class="wrapc">${escHtml(displayPlan(e))}</td>
      <td class="wrapc">${escHtml(e.allergies)}</td>
      <td class="wrapc">${escHtml(e.removables)}</td>
      <td class="wrapc">${escHtml(e.dislikemeals)}</td>
      <td class="wrapc">${escHtml(customNotes(e))}</td>
      <td>${escHtml(e.packnote)}</td>
      <td class="mono">${escHtml(e.duration)}</td>
      <td class="mono">${escHtml(e.start)}</td>
      <td class="mono">${escHtml(due)}${due && e.dueBase && due!==e.dueBase ? `<div class="foot-note">was ${escHtml(e.dueBase)}</div>` : ''}</td>
      <td><label class="sw" title="Freeze switch"><input type="checkbox" data-act="freeze" data-id="${e.id}" ${isFrozenNow(e)?'checked':''} aria-label="Frozen, ${escHtml(e.name)}"><span></span></label></td>
      <td><span class="dayticks">${WEEKDAYS.map(d=>`<button type="button" data-act="day" data-day="${d}" data-id="${e.id}" aria-pressed="${(e.daysOff||[]).includes(d)}" title="${d} off">${DAY_LETTER[d]}</button>`).join('')}</span></td>
      <td class="mono">${rem===''?'':rem}</td>
      <td class="wrapc">${escHtml(locText(e.locA))}</td>
      <td class="wrapc">${escHtml(locText(e.locB))}</td>
      <td><div class="act"><button type="button" data-act="edit" data-id="${e.id}">Edit</button>
        <button type="button" data-act="expire" data-id="${e.id}">${/EXPIR/i.test(e.status)?'Reactivate':'Expire'}</button>
        <button type="button" class="danger" data-act="del" data-id="${e.id}">${pendingDelete===e.id?'Confirm delete':'Delete'}</button></div></td>
    </tr>`;
  });
  if(!view.length) html += `<tr><td colspan="20" class="empty">No customers match.</td></tr>`;
  const t = rosterTotals(today);
  const foot = (lab, o)=>`<tr><td colspan="3">${lab}</td><td colspan="2" class="mono">Keto ${o.keto}</td><td class="mono">Macro ${o.macro}</td><td class="mono" colspan="2">Simple Diet ${o.simple}</td><td class="mono" colspan="2">${o.app?('App '+o.app):''}</td><td class="mono" colspan="10">Total ${o.total}</td></tr>`;
  html += `</tbody><tfoot>${foot('All customers', t.all)}${foot('Active today', t.act)}${foot('Frozen now', t.frz)}</tfoot></table></div>
    <p class="foot-note" style="margin-top:8px;">Showing ${view.length} of ${roster.length}. Remaining days count serving days after today; off days and frozen days do not use up the package.</p>`;
  host.innerHTML = html;
}
(function wireGrid(){
  const host = $('#gridHost'); if(!host) return;
  host.addEventListener('click', ev=>{
    const b = ev.target.closest ? ev.target.closest('button[data-act]') : null; if(!b) return;
    const e = roster.find(x=>x.id===b.dataset.id); if(!e) return;
    const act = b.dataset.act;
    if(act==='edit'){ pendingDelete = null; openReg(e); return; }
    if(act==='day'){
      const d = b.dataset.day, set = new Set(e.daysOff||[]);
      if(set.has(d)) set.delete(d); else set.add(d);
      e.daysOff = WEEKDAYS.filter(x=>set.has(x));
      if(e.dueAuto) e.dueBase = computeDueBase(e);
    }
    if(act==='expire'){ e.status = /EXPIR/i.test(e.status) ? 'ACTIVE' : 'EXPIRED'; }
    if(act==='del'){
      if(pendingDelete!==e.id){ pendingDelete = e.id; renderGrid(); return; }
      roster = roster.filter(x=>x.id!==e.id); pendingDelete = null;
    }
    saveRoster();
  });
  host.addEventListener('change', ev=>{
    const i = ev.target;
    if(!i || i.dataset.act!=='freeze') return;
    const e = roster.find(x=>x.id===i.dataset.id); if(!e) return;
    setFreeze(e, i.checked, todayISO());
    saveRoster();
  });
  ['#gSearch','#gHk','#gStatus'].forEach(s=>{ $(s).addEventListener('input', renderGrid); $(s).addEventListener('change', renderGrid); });
})();

/* ======================= registration form ======================= */
let regEditing = null, regPeriods = [], regDueTouched = false;
const LOCF = [['zone','Zone','text'],['street','Street','text'],['building','Building','text'],['fromDay','From day','day'],['toDay','To day','day'],['fromDate','From date','date'],['toDate','To date','date']];
(function buildRegStatic(){
  if(!$('#r_locs')) return;
  $('#r_locs').innerHTML = ['a','b'].map(k=>`<div class="locrow"><div class="ttl">Location ${k.toUpperCase()}</div>${LOCF.map(f=>`<div class="field"><label for="r_l${k}_${f[0]}">${f[1]}</label>${
    f[2]==='day' ? `<select id="r_l${k}_${f[0]}"><option value=""></option>${WEEKDAYS.map(d=>`<option>${d}</option>`).join('')}</select>` : `<input type="${f[2]==='date'?'date':'text'}" id="r_l${k}_${f[0]}">`}</div>`).join('')}</div>`).join('');
  $('#r_days').innerHTML = WEEKDAYS.map(d=>`<label><input type="checkbox" name="dayoff" value="${d}"> ${d}</label>`).join('');
})();
const rv = id=> str($('#'+id).value);
const setv = (id,v)=>{ $('#'+id).value = v==null ? '' : v; };
function nextBox(hk){
  const pre = {keto:'K', macro:'M', simple:'SD', app:'A'}[hk] || 'M';
  let max = 0;
  roster.forEach(c=>{ const m = new RegExp('^'+pre+'(\\d+)$').exec(c.box); if(m) max = Math.max(max, +m[1]); });
  return pre + (max+1);
}
function readReg(){
  const base = regEditing ? (roster.find(x=>x.id===regEditing)||{}) : {};
  const hk = rv('r_hk');
  const planSel = $('#r_plan').value;
  const e = Object.assign({}, base, {
    id: regEditing || uid(), hk, status: rv('r_status') || 'ACTIVE', contact: rv('r_contact'), box: rv('r_box').toUpperCase().replace(/\s+/g,''), name: rv('r_name'),
    plan: planSel==='__raw__' ? '' : planSel, mealplan: planSel==='__raw__' ? (base.mealplan||'') : '',
    gramType: (hk==='keto'||hk==='app') ? 'classic' : rv('r_gram'),
    grams: {lp:rv('r_lp'), lc:rv('r_lc'), dp:rv('r_dp'), dc:rv('r_dc')},
    allergies: rv('r_allergies'), removables: rv('r_removables'), dislikemeals: rv('r_dislikemeals'), notes: $('#r_notes').value.trim(), packnote: rv('r_packnote'),
    duration: rv('r_duration')==='' ? '' : +rv('r_duration'), start: rv('r_start'), dueBase: rv('r_due'),
    daysOff: $$('#r_days input').filter(i=>i.checked).map(i=>i.value), freezePeriods: regPeriods.map(p=>({from:p.from, to:p.to})),
    locA: {}, locB: {}
  });
  ['a','b'].forEach(k=>{ const o = {}; LOCF.forEach(f=>{ o[f[0]] = rv('r_l'+k+'_'+f[0]); }); e[k==='a'?'locA':'locB'] = o; });
  if(!e.xNote) e.xNote = ''; if(!e.xFreeze) e.xFreeze = '';
  return e;
}
function syncRegUI(){
  const hk = rv('r_hk');
  $('#r_gram').disabled = (hk==='keto' || hk==='app');
  $('#r_customBox').hidden = (hk==='keto' || hk==='app') || rv('r_gram')!=='custom';
}
function updateRegDue(){
  const e = readReg(), today = todayISO();
  if(!regDueTouched){ const d = computeDueBase(e); setv('r_due', d); e.dueBase = d; }
  setv('r_dueEff', e.dueBase ? dueEff(e,today) : '');
  const rem = remainingDays(e,today);
  setv('r_remaining', rem===''?'':rem);
  $('#r_freezeOn').checked = isFrozenNow(e);
}
function renderRegPeriods(){
  $('#r_periods').innerHTML = regPeriods.map((p,i)=>`<span class="chip">${escHtml(p.from)} → ${p.to?escHtml(p.to):'ongoing'}<button type="button" data-pi="${i}" aria-label="Remove freeze period">×</button></span>`).join('');
  $$('#r_periods button').forEach(b=> b.addEventListener('click', ()=>{ regPeriods.splice(+b.dataset.pi,1); renderRegPeriods(); updateRegDue(); }));
  updateRegDue();
}
function openReg(e){
  regEditing = e ? e.id : null;
  regDueTouched = e ? !e.dueAuto : false;
  regPeriods = e ? (e.freezePeriods||[]).map(p=>({from:p.from,to:p.to})) : [];
  $('#regTitle').textContent = e ? 'Edit '+e.name : 'Register customer';
  const hkSel = $('#r_hk');
  if(e && e.hk==='app' && !$$('#r_hk option').some(o=>o.value==='app')){ const o = document.createElement('option'); o.value='app'; o.textContent='App'; hkSel.appendChild(o); }
  setv('r_hk', e ? e.hk : 'macro');
  setv('r_status', e ? (/EXPIR/i.test(e.status) ? 'EXPIRED' : 'ACTIVE') : 'ACTIVE');
  setv('r_contact', e&&e.contact); setv('r_name', e&&e.name);
  setv('r_box', e ? e.box : nextBox('macro'));
  $('#r_plan').innerHTML = (e && !e.plan && e.mealplan ? `<option value="__raw__">As in the sheet: ${escHtml(e.mealplan)}</option>` : '') + PLANS.map(p=>`<option value="${escHtml(p[0])}">${escHtml(p[0])}</option>`).join('');
  setv('r_plan', e ? (e.plan || '__raw__') : PLANS[0][0]);
  setv('r_gram', e ? e.gramType : 'classic');
  const g = (e&&e.grams)||{}; setv('r_lp',g.lp); setv('r_lc',g.lc); setv('r_dp',g.dp); setv('r_dc',g.dc);
  setv('r_allergies', e&&e.allergies); setv('r_removables', e&&e.removables); setv('r_dislikemeals', e&&e.dislikemeals);
  setv('r_notes', e&&e.notes); setv('r_packnote', e&&e.packnote);
  setv('r_duration', e ? e.duration : 24); setv('r_start', e ? e.start : todayISO()); setv('r_due', e ? e.dueBase : '');
  setv('r_fzFrom',''); setv('r_fzTo','');
  $$('#r_days input').forEach(i=>{ i.checked = !!(e && (e.daysOff||[]).includes(i.value)); });
  ['a','b'].forEach(k=>{ const l = (e && (k==='a'?e.locA:e.locB)) || {}; LOCF.forEach(f=> setv('r_l'+k+'_'+f[0], l[f[0]])); });
  $('#regError').textContent = '';
  syncRegUI(); renderRegPeriods();
  $('#regCard').hidden = false;
  $('#r_name').focus();
  if(window.scrollTo) window.scrollTo(0,0);
}
function closeReg(){ $('#regCard').hidden = true; regEditing = null; }
(function wireReg(){
  const form = $('#regForm'); if(!form) return;
  $('#regOpen').addEventListener('click', ()=> openReg(null));
  $('#regCancel').addEventListener('click', closeReg);
  let boxTouched = false;
  $('#r_box').addEventListener('input', ()=>{ boxTouched = true; });
  $('#r_hk').addEventListener('change', ()=>{ if(!regEditing && !boxTouched) setv('r_box', nextBox(rv('r_hk'))); syncRegUI(); updateRegDue(); });
  $('#r_gram').addEventListener('change', syncRegUI);
  ['r_duration','r_start'].forEach(id=> $('#'+id).addEventListener('input', updateRegDue));
  $('#r_days').addEventListener('change', updateRegDue);
  $('#r_due').addEventListener('input', ()=>{ regDueTouched = rv('r_due')!==''; updateRegDue(); });
  $('#r_freezeOn').addEventListener('change', ()=>{
    const tmp = {freezePeriods: regPeriods}; setFreeze(tmp, $('#r_freezeOn').checked, todayISO()); regPeriods = tmp.freezePeriods; renderRegPeriods();
  });
  $('#r_fzAdd').addEventListener('click', ()=>{
    const a = rv('r_fzFrom'), b = rv('r_fzTo') || a; if(!a) return;
    regPeriods.push({from:a, to: b < a ? a : b}); setv('r_fzFrom',''); setv('r_fzTo',''); renderRegPeriods();
  });
  form.addEventListener('submit', ev=>{
    ev.preventDefault();
    const e = readReg();
    let err = '';
    if(!e.name) err = 'Enter the customer name.';
    else if(!e.contact) err = 'Enter the contact number.';
    else if(!e.box) err = 'Enter a box number.';
    else if(roster.some(c=>c.box===e.box && c.id!==e.id)) err = 'Box '+e.box+' is already used by another customer.';
    else if(!e.start) err = 'Choose the joining date.';
    else if(!e.dueBase) err = 'Enter the package duration, or type a due date.';
    else if(e.dueBase < e.start) err = 'The due date is before the joining date.';
    else if(e.gramType==='custom' && !(e.grams.lp||e.grams.lc||e.grams.dp||e.grams.dc)) err = 'Enter the customized grams for lunch and dinner, or pick Classic or HP.';
    if(err){ $('#regError').textContent = err; return; }
    e.dueAuto = !regDueTouched;
    const i = roster.findIndex(x=>x.id===e.id);
    if(i>=0) roster[i] = e; else roster.push(e);
    closeReg(); saveRoster();
  });
})();

/* ======================= counts, manual clients, orders ======================= */
const COUNT_SLOTS = [['breakfast','Breakfast'],['lunch','Lunch'],['dinner','Dinner'],['amsnack','AM snack'],['pmsnack','PM snack']];
function countsMatrix(sheet){
  const cell = (L, slot)=>{
    let n = 0, d = 0; const add = it=>{ n++; if(it.nos.length) d++; };
    if(slot==='special') Object.values(L.special).forEach(o=>o.items.forEach(add));
    else { const S = L.slots.find(s=>s.slot===slot); S.rows.forEach(r=>r.items.forEach(add)); Object.values(S.others).forEach(o=>o.items.forEach(add)); }
    return {n,d};
  };
  const mkRow = (label, slots)=>{ const cells = sheet.lines.map(L=> slots.map(s=>cell(L,s)).reduce((a,c)=>({n:a.n+c.n,d:a.d+c.d}),{n:0,d:0})); return {label, cells, n:cells.reduce((a,c)=>a+c.n,0), d:cells.reduce((a,c)=>a+c.d,0)}; };
  const rows = COUNT_SLOTS.map(([s,l])=>mkRow(l,[s]));
  const snacks = mkRow('Snacks (AM + PM)', ['amsnack','pmsnack']);
  const special = mkRow('Special add-ons', ['special']);
  const total = mkRow('Total meals', COUNT_SLOTS.map(x=>x[0]).concat(['special']));
  return {rows:rows.concat([snacks]).concat(special.n?[special]:[]), total};
}
function renderCounts(sheet){
  const m = countsMatrix(sheet), host = $('#countsTable'); if(!host) return;
  const line = r=>`<tr${r===m.total?' class="tot"':''}><td>${escHtml(r.label)}</td>${r.cells.map(c=>`<td class="num">${c.n}</td>`).join('')}<td class="num">${r.n}</td><td class="num">${r.d||''}</td></tr>`;
  host.innerHTML = `<div class="tbl-wrap cnt"><table style="min-width:420px;"><thead><tr><th>Meal</th>${sheet.lines.map(L=>`<th class="num">${escHtml(L.label)}</th>`).join('')}<th class="num">Total</th><th class="num">With dislikes</th></tr></thead><tbody>${m.rows.map(line).join('')}${line(m.total)}</tbody></table></div>
    <p class="foot-note" style="margin-top:8px;">The rows of each meal, with the quantity and the number with dislikes for every dish, are in the sheet below.</p>`;
}
function manualRows(){
  const sh = runResults && runResults.sheet; if(!sh) return [['Nothing yet']];
  const rows = [['Box','HK','Customer name','Contact','Meal plan package','Why it is here','Pack note','Customer notes','Allergies','Dislikes (ingredients)','Dislike meals']];
  sh.remaining.forEach(x=>{ const c = x.raw || {};
    rows.push([x.box, LINE_LABEL[x.diet]||'', x.name, c.contact||'', str(c.mealplan), x.reasons.join('; '), str(c.packnote), str(c.custnotes), str(c.allergies), str(c.removables||c.dislike), str(c.dislikemeals)]); });
  return rows;
}
function renderManual(){
  const host = $('#manualTable'); if(!host) return;
  const rows = manualRows();
  $('#manualCount').textContent = String(Math.max(0, rows.length-1));
  if(rows.length<2){ host.innerHTML = '<div class="empty">Nobody needs doing by hand for this date.</div>'; return; }
  host.innerHTML = `<div class="tbl-wrap"><table><thead><tr>${rows[0].map(h=>`<th>${escHtml(h)}</th>`).join('')}</tr></thead><tbody>${
    rows.slice(1).map(r=>`<tr>${r.map((c,i)=>`<td class="${i===0?'mono':''}" ${i===5?'':''}>${escHtml(c)}</td>`).join('')}</tr>`).join('')}</tbody></table></div>`;
}
const PROT_LABEL = {chicken:'Chicken',beef:'Beef / meat',salmon:'Salmon',shrimp:'Shrimp',tuna:'Tuna',fish:'White fish',turkey:'Turkey',egg:'Eggs',halloumi:'Halloumi'};
function parseGramStr(g){ const m = /^(\d+)(?:\+(\d+))?/.exec(g||''); return m ? {p:+m[1], c:m[2]?+m[2]:0} : null; }
function lineItems(L){
  const out = [];
  L.slots.forEach(S=>{ S.rows.forEach(r=>r.items.forEach(i=>out.push(i))); Object.values(S.others).forEach(o=>o.items.forEach(i=>out.push(i))); });
  Object.values(L.special).forEach(o=>o.items.forEach(i=>out.push(i)));
  return out;
}
function orderTotals(sheet){
  const prot = {}, carb = {};
  const bump = (map, name, it, g)=>{
    const o = map[name] || (map[name] = {portions:0, withG:0, grams:0, dishes:{}});
    o.portions++; if(g){ o.withG++; o.grams += g; } o.dishes[it.dish] = (o.dishes[it.dish]||0)+1;
  };
  sheet.lines.forEach(L=> lineItems(L).forEach(it=>{
    const g = parseGramStr(it.og), p = proteinOf(it.dish);
    if(p || (g && g.p)) bump(prot, p ? PROT_LABEL[p] : 'Other (no protein named in the dish)', it, g && g.p);
    const cs = carbsIn(it.dish);
    cs.forEach(c=> bump(carb, c, it, g && g.c ? g.c/cs.length : 0));
  }));
  return {prot, carb};
}
function orderBuffer(){ const v = parseFloat($('#ordBuffer') && $('#ordBuffer').value); return isNaN(v) ? 0 : Math.max(0, v); }
function orderRows(){
  if(!runResults || !runResults.sheet) return [['Nothing yet']];
  const t = orderTotals(runResults.sheet), b = orderBuffer();
  const rows = [['Type','Item','Portions','Portions with a weight','Total grams','Total kg','Kg with buffer '+b+'%','Dishes']];
  const add = (type, map)=> Object.keys(map).sort((x,y)=>map[y].grams-map[x].grams || map[y].portions-map[x].portions).forEach(k=>{
    const o = map[k];
    rows.push([type, k, o.portions, o.withG, Math.round(o.grams), +(o.grams/1000).toFixed(2), +(o.grams*(1+b/100)/1000).toFixed(2), Object.entries(o.dishes).sort((x,y)=>y[1]-x[1]).map(d=>d[0]+' x'+d[1]).join('; ')]);
  });
  add('Protein', t.prot); add('Carb', t.carb);
  return rows;
}
function renderOrders(){
  const host = $('#ordersTable'); if(!host) return;
  const rows = orderRows();
  if(rows.length<2){ host.innerHTML = '<div class="empty">No meals to order for this date.</div>'; return; }
  const block = type=>{
    const r = rows.slice(1).filter(x=>x[0]===type);
    if(!r.length) return '';
    return `<h3 style="margin:14px 0 6px;">${type==='Protein'?'Proteins':'Carbs'}</h3><div class="tbl-wrap cnt"><table><thead><tr><th>Item</th><th class="num">Portions</th><th class="num">With a weight</th><th class="num">Kg</th><th class="num">Kg with buffer</th><th>Dishes</th></tr></thead><tbody>${
      r.map(x=>`<tr><td><strong>${escHtml(x[1])}</strong></td><td class="num">${x[2]}</td><td class="num">${x[3]}</td><td class="num">${x[5]}</td><td class="num">${x[6]}</td><td class="foot-note" style="font-family:inherit;text-align:left;white-space:normal;">${escHtml(x[7])}</td></tr>`).join('')}</tbody></table></div>`;
  };
  host.innerHTML = block('Protein') + block('Carb');
}

/* ======================= meal sheet + production rows for the workbook ======================= */
function mealRows(){
  const rows = [['Line','Meal','Main dish','Option 1','Option 2']];
  [['keto','Keto'],['macro','Macro'],['simple','Simple Diet']].forEach(([k,lab])=>{
    COUNT_SLOTS.forEach(([s,sl])=>{
      const p = poolFor(k, s);
      rows.push([lab, sl, (p.find(c=>c.flag==='main')||{}).name||'', (p.find(c=>c.flag==='opt1')||{}).name||'', (p.find(c=>c.flag==='opt2')||{}).name||'']);
    });
  });
  return rows;
}
function productionRows(){
  const sh = runResults && runResults.sheet; if(!sh) return [['Nothing yet']];
  const m = countsMatrix(sh);
  const rows = [['PRODUCTION SHEET', runResults.weekday+' '+runResults.prodDate], [], ['Meal'].concat(sh.lines.map(L=>L.label), ['Total','With dislikes'])];
  m.rows.concat([m.total]).forEach(r=> rows.push([r.label].concat(r.cells.map(c=>c.n), [r.n, r.d||''])));
  rows.push([]); rows.push(['Line','Row','Meal','Qty','With dislikes','Boxes (grammage--box; dislikes last)']);
  sh.lines.forEach(L=> sheetRows(L).forEach(r=> rows.push([L.label, r.label, r.dish, r.items.length||'', r.items.filter(i=>i.nos.length).length||'', groupsText(r.items)])));
  rows.push([]); rows.push(['BOXES TO SELECT BY HAND', sh.remaining.length, sh.remaining.map(x=>x.box).join(' / ')]);
  return rows;
}

/* ======================= workbook export ======================= */
const SHEET_HEAD = ['HK','STATUS','CONTACT #','BOX #','CUSTOMER NAME','MEAL PLAN PACKAGE','ALLERGIES','DISLIKE (INGREDIENTS)','DISLIKE MEALS','CUSTOMER NOTES','PACK. NOTE','PACKAGE DURATION','JOINING DATE','DUE DATE','OFF DAYS','REMAINING DAYS',''];
function locHead(k){ return ['ZONE','STREET','BUILDING','FROM DAY','TO DAY','FROM DATE','TO DATE'].map(x=>'LOCATION '+k+' '+x); }
const SHEET_HEAD_ALL = SHEET_HEAD.concat(locHead('A'), locHead('B'), ['FREEZE (ON/OFF)','FREEZE PERIODS','FROZEN DAYS','TODAY','PLAN (APP)','NOTE','FREEZE DATE (OLD SHEET)','DUE DATE BEFORE FREEZES']);
function customerSheetRows(){
  const today = todayISO(), rows = [SHEET_HEAD_ALL];
  const locCells = l=>{ l = l||{}; return [l.zone||'',l.street||'',l.building||'',l.fromDay||'',l.toDay||'',l.fromDate||'',l.toDate||'']; };
  sortedRoster().forEach(e=>{
    const rem = remainingDays(e,today), raw = str(e.status).toUpperCase() || 'ACTIVE';
    const status = /EXPIR/.test(raw) ? 'EXPIRED' : (isFrozenNow(e) ? 'FROZEN' : raw);
    const fd = new Set(); (e.freezePeriods||[]).forEach(p=>{ let d=p.from, g=0; const to = p.to==null ? (today>p.from?today:p.from) : p.to; while(d<=to && g++<800){ fd.add(d); d = addDaysISO(d,1); } });
    rows.push([PKG_LABEL[e.hk]||'', status, e.contact, e.box, e.name, displayPlan(e), e.allergies, e.removables, e.dislikemeals, customNotes(e), e.packnote,
      e.duration, e.start, dueEff(e,today), (e.daysOff||[]).join(', '), rem, ''].concat(locCells(e.locA), locCells(e.locB),
      [isFrozenNow(e)?'ON':'OFF', periodsText(e), fd.size, today, e.plan||'', e.xNote||'', e.xFreeze||'', e.dueBase||'']));
  });
  const t = rosterTotals(today);
  rows.push([]);
  rows.push(['TOTALS','ALL CUSTOMERS','ACTIVE','FROZEN']);
  [['KETO','keto'],['MACRO','macro'],['SIMPLE DIET','simple'],['APP','app'],['TOTAL','total']].forEach(([l,k])=> rows.push([l, t.all[k], t.act[k], t.frz[k]]));
  return rows;
}
function buildWorkbook(){
  runEngine();
  const wb = XLSX.utils.book_new();
  const add = (name, rows, widths)=>{ const ws = XLSX.utils.aoa_to_sheet(rows); if(widths) ws['!cols'] = widths.map(w=>({wch:w})); XLSX.utils.book_append_sheet(wb, ws, name); };
  add('Customer Data', customerSheetRows(), [12,10,12,9,24,34,16,24,22,34,14,10,12,12,18,10,3].concat(Array(30).fill(16)));
  add('Manual Clients', manualRows(), [9,12,24,12,32,50,14,34,16,22,22]);
  add('Protein & Carb Orders', orderRows(), [10,34,10,12,12,10,14,60]);
  add('Production Sheet', productionRows(), [22,18,40,8,14,70]);
  add('Meal Sheet', mealRows(), [12,12,36,36,36]);
  return wb;
}
$('#wbExport').dataset.label = 'Export workbook (Excel)';
$('#wbExport').addEventListener('click', e=>{
  if(!roster.length){ flash(e.target, 'Nothing to export'); return; }
  const wb = buildWorkbook(), data = XLSX.write(wb, {bookType:'xlsx', type:'array'});
  deliver('kitchen-run-'+prodDateInput.value+'.xlsx', data, e.target, toCSV(customerSheetRows()));
});
$('#exportManual').dataset.label = 'Copy CSV';
$('#exportManual').addEventListener('click', e=> deliver('manual-clients-'+prodDateInput.value+'.csv', toCSV(manualRows()), e.target));
$('#exportOrders').dataset.label = 'Copy CSV';
$('#exportOrders').addEventListener('click', e=> deliver('protein-carb-orders-'+prodDateInput.value+'.csv', toCSV(orderRows()), e.target));
$('#ordBuffer').value = lsGet('kr_buffer', 0);
$('#ordBuffer').addEventListener('input', ()=>{ lsSet('kr_buffer', orderBuffer()); renderOrders(); });

/* ======================= start-up ======================= */
$('#rulesBody').innerHTML = `
  <p><strong>Who is produced.</strong> Status Active, not frozen on the date, production date between the joining date and the due date, not an off day, Simple Diet closed on Fridays. Food-schedule customers are never produced; their boxes go to the Manual clients sheet.</p>
  <p><strong>Meals in a plan.</strong> 3 meals + 2 snacks: breakfast, lunch, dinner, AM, PM. 2 meals: lunch and dinner. 3 meals with no breakfast: 2 lunches + 1 dinner or 1 lunch + 2 dinners, different dishes. 3 meals + 1 snack adds one snack. Business plan: lunch plus 1 snack. A plan the app cannot read goes to Manual clients.</p>
  <p><strong>Grammage.</strong> Macro Classic 100+100 lunch, 100 dinner; HP 150+100 lunch, 150 dinner. Customer notes with their own grams replace these. Keto and Simple Diet show no grammage (Simple Diet is counted only), unless the customer has their own grams.</p>
  <p><strong>Dislikes.</strong> Ingredient dislikes stay on the dish and print beside the box. Dislike meals and allergies mean a different dish; if none is safe, the box goes to Manual clients. Customers with dislikes are listed last in each row.</p>
  <p><strong>Variety.</strong> Customers are spread across the main dish and the options, and a customer does not get the same dish or protein twice in a day when another safe dish exists.</p>
  <p><strong>Days and remaining days.</strong> The due date is the day the last meal of the package is served. Off days and frozen days do not use up the package, so each frozen serving day moves the due date one serving day on. Remaining days are the serving days left after today.</p>`;

/* ---------- home, logo ---------- */
function setNum(el, v){
  const to = +v || 0, from = +el.dataset.v || 0; el.dataset.v = to;
  if(from===to || window.matchMedia && matchMedia('(prefers-reduced-motion: reduce)').matches){ el.textContent = to; return; }
  const t0 = performance.now(), dur = 600;
  (function step(t){ const k = Math.min(1,(t-t0)/dur), e = 1-Math.pow(1-k,3); el.textContent = Math.round(from+(to-from)*e); if(k<1) requestAnimationFrame(step); })(t0);
}
function renderHome(){
  if(!$('#panel-home')) return;
  const HK_LABEL = {keto:'Keto', macro:'Macro', simple:'Simple Diet', app:'App'};
  const today = todayISO(), tot = rosterTotals(today);
  setNum($('#kActive'), tot.act.total); setNum($('#kFrozen'), tot.frz.total);
  setNum($('#kAll'), tot.all.total);
  setNum($('#kExpired'), roster.filter(e=>statusInfo(e,today)[0]==='Expired').length);
  const pd = prodDateInput.value || today, d = new Date(pd+'T00:00:00');
  const mon = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];
  $('#tWeek').textContent = isNaN(d) ? '—' : WEEKDAYS[d.getDay()].toUpperCase();
  $('#tDate').textContent = isNaN(d) ? '—' : d.getDate()+' '+mon[d.getMonth()]+' '+d.getFullYear();
  $('#tBoxes').textContent = $('#eligCount').textContent;
  $('#tActive').textContent = tot.act.total; $('#tFrozen').textContent = tot.frz.total;
  const max = Math.max(1, tot.act.keto, tot.act.macro, tot.act.simple, tot.act.app);
  $('#hMix').innerHTML = ['keto','macro','simple','app'].map(k=>
    `<button class="h-row" data-go="data" data-hk="${k}"><span>${HK_LABEL[k]}</span><span class="h-bar"><i data-w="${Math.round(tot.act[k]/max*100)}"></i></span><span class="v">${tot.act[k]}</span></button>`).join('');
  requestAnimationFrame(()=> $$('#hMix i').forEach(i=>{ i.style.width = i.dataset.w+'%'; }));
  const soon = roster.map(e=>({e, st:statusInfo(e,today)[0], rem:remainingDays(e,today)}))
    .filter(x=>x.st==='Active' && x.rem!=='' && x.rem<=3).sort((a,b)=>a.rem-b.rem).slice(0,8);
  $('#hSoon').innerHTML = soon.length ? soon.map(x=>
    `<button data-go="data" data-q="${escHtml(x.e.box||x.e.name)}"><span><b>${escHtml(x.e.name)}</b> <small>${escHtml(x.e.box||'')} · ${HK_LABEL[x.e.hk]||''}</small></span><small>${x.rem===0?'last day':x.rem+(x.rem===1?' day':' days')+' left'}</small></button>`).join('')
    : `<div class="empty">${roster.length ? 'Nobody ends in the next 3 days.' : 'Load a customer workbook or register a customer to see who is ending soon.'}</div>`;
}
function applyLogo(){
  $$('[data-logo]').forEach(el=>{
    el.innerHTML = '';
    const im = document.createElement('img');
    im.src = 'logo.svg'; im.alt = 'Anino Cloud';
    el.appendChild(im);
  });
}
function shiftProdDate(n){
  const base = prodDateInput.value || todayISO();
  prodDateInput.value = n===0 ? todayISO() : addDaysISO(base,n);
  prodDateInput.dispatchEvent(new Event('change'));
}
$('#tPrev').addEventListener('click', ()=>shiftProdDate(-1));
$('#tNext').addEventListener('click', ()=>shiftProdDate(1));
$('#tToday').addEventListener('click', ()=>shiftProdDate(0));
$('#hBuild').addEventListener('click', ()=>{ showTab('sheet'); window.scrollTo(0,0); });
$('#hReg').addEventListener('click', ()=>{ showTab('data'); const b = $('#regOpen'); if(b) b.click(); });
$('#panel-home').addEventListener('click', e=>{
  const t = e.target.closest('[data-go]'); if(!t) return;
  const go = t.dataset.go;
  if(go==='data'){
    $('#gSearch').value = t.dataset.q || ''; $('#gHk').value = t.dataset.hk || 'all'; $('#gStatus').value = t.dataset.status || 'all';
    ['input','change'].forEach(n=>{ $('#gSearch').dispatchEvent(new Event(n)); $('#gHk').dispatchEvent(new Event(n)); $('#gStatus').dispatchEvent(new Event(n)); });
    renderGrid();
  }
  showTab(go); window.scrollTo(0,0);
});
initRoster();
applyLogo();
refreshDateReadout();
renderCustStatus();
renderGrid();
refreshEligCount();
renderHome();

})();
