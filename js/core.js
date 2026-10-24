/* Pure parsing + calculation. No DOM access, so it runs in the browser and under `node --test`. */
(function (root) {
"use strict";

const NON_PEN = ["PREMIX","2PRMIX","BOVAR","TOTAL",""];
const PREMIX_CODES = ["PREMIX","2PRMIX"];
const HEADERS = ["Producer Ingredient Name","Pounds/Head/Day","Pounds/Month","Tons/Month","Dry Matter/As Fed","Cow Type"];
const LB_PER_TON = 2000;
const NOP_IN_BOVAER = 0.10;   // Bovaer 10 label: 10% 3-NOP, as-fed basis
const DOSE_TARGET = 60;       // mg 3-NOP per kg DM intake
const LABEL_MIN = 60, LABEL_MAX = 63;   // Bovaer Dose Calculator "on target" band
const MG_PER_LB = 453592.37, KG_PER_LB = 0.45359237;

class UserError extends Error {}

/* ---------- CSV ---------- */
function parseCSV(text){
  text = text.replace(/^\uFEFF/, "");
  const rows = []; let row = [], field = "", q = false;
  for (let i=0;i<text.length;i++){
    const c = text[i];
    if (q){ if (c === '"'){ if (text[i+1] === '"'){ field+='"'; i++; } else q=false; } else field += c; }
    else { if (c === '"') q = true;
      else if (c === ',') { row.push(field); field=""; }
      else if (c === '\n'){ row.push(field); rows.push(row); row=[]; field=""; }
      else if (c === '\r'){}
      else field += c; }
  }
  if (field.length || row.length){ row.push(field); rows.push(row); }
  return rows;
}
function toRecords(grid){
  const head = grid[0].map(h => String(h).trim()); const recs = [];
  for (let i=1;i<grid.length;i++){
    if (grid[i].length === 1 && grid[i][0] === "") continue;
    const o = {}; head.forEach((h,j) => o[h] = (grid[i][j] !== undefined ? String(grid[i][j]).trim() : "")); recs.push(o);
  }
  return recs;
}

/* ---------- TMR Tracker ingredient report, native Excel layout ---------- */
// Fixed columns: 0=code, 1=desc, 4="SubTotal"/"Total" marker, 7=Actual Wt, 8=Dry Weight, 9=Units (or head count on a Pen row)
function rowsToRecords(rows){
  const S = v => (v===undefined||v===null) ? "" : String(v).trim();
  // Column positions vary between exports — detect them from the header row labels.
  let cCode=0, cDesc=1, cAct=7, cDry=8, cUnit=9;
  for (const r of rows){
    const idx = {}; r.forEach((c,j)=>{ const s=S(c); if(s && !(s in idx)) idx[s]=j; });
    if (("Display Name" in idx) && ("Actual Wt." in idx)){
      cCode = idx["Display Name"];
      cDesc = ("Description" in idx) ? idx["Description"] : cCode+1;
      cAct  = idx["Actual Wt."];
      cDry  = ("Dry Weight" in idx) ? idx["Dry Weight"] : cAct+1;
      cUnit = ("Units" in idx) ? idx["Units"] : cDry+1;
      break;
    }
  }
  const has = (r,val) => r.some(c=>S(c)===val);
  const firstNum = r => { for(const c of r){ if(typeof c==="number" && isFinite(c)) return c;
      const s=S(c); if(s && /^[\d,]+(\.\d+)?$/.test(s)) return parseFloat(s.replace(/,/g,"")); } return ""; };
  const num = v => isFinite(parseFloat(v));

  const recs = []; let pen="", heads="", grand=false;
  for (const r of rows){
    const penCell = r.find(c=>S(c).startsWith("Pen :"));
    if (penCell !== undefined){ pen = S(penCell).slice(5).split(",")[0].trim(); heads = firstNum(r); grand=false; continue; }
    if (has(r,"Display Name")) continue;                         // repeating page header
    if (has(r,"Totals")){ grand = true; pen = "TOTAL"; continue; }
    if (has(r,"Total")){                                          // final grand-total line
      recs.push({Pen:"TOTAL", AvgHeads:"", RowType:"GrandTotalLine", DisplayName:"", Description:"",
                 ActualWt_Ton:S(r[cAct]), DryWt_Ton:S(r[cDry]), Units:S(r[cUnit])}); continue; }
    if (has(r,"SubTotal")){
      recs.push({Pen:pen, AvgHeads:heads, RowType:"SubTotal", DisplayName:"", Description:"",
                 ActualWt_Ton:S(r[cAct]), DryWt_Ton:S(r[cDry]), Units:S(r[cUnit])}); continue; }
    const code = S(r[cCode]);
    if (code && (num(r[cAct]) || num(r[cDry]))){
      recs.push({Pen:(grand?"TOTAL":pen), AvgHeads:(grand?"":heads), RowType:(grand?"GrandTotal":"Ingredient"),
                 DisplayName:code, Description:S(r[cDesc]), ActualWt_Ton:S(r[cAct]), DryWt_Ton:S(r[cDry]), Units:S(r[cUnit])});
    }
  }
  return recs;
}

/* ---------- TapFeed Load report (CSV, one row per ingredient load or pen delivery) ---------- */
const codeOf = s => { const m = String(s).match(/\(([^()]*)\)\s*$/); return m ? m[1].trim() : String(s).trim(); };
const nameOf = s => String(s).replace(/\s*\([^()]*\)\s*$/, "").trim();
const usDate = iso => { const [y,m,d] = iso.split("-").map(Number); return `${m}/${d}/${y}`; };

// Single-month method, checkable from the report alone: only the last month in the file is counted, and the Bovaer %
// and Bovaer premix % come from the batches made in that month. Bovaer fed = premix fed (as fed) x Bovaer premix %
// x Bovaer %. Lactating = rations that contain a Bovaer-carrying premix (PREMIX / 2PRMIX).
function fromLoadReport(grid){
  const head = grid[0].map(h => String(h).trim());
  const cBatch = head.indexOf("BatchId"), cRecipe = head.indexOf("Recipe"), cIng = head.indexOf("Ingredient"),
        cDate = head.indexOf("Load/Delivery date and time"), cDry = head.indexOf("Dry weight"),
        cAct = head.indexOf("Actual weight");
  const S = v => (v===undefined||v===null) ? "" : String(v).trim();
  const lbs = v => parseFloat(S(v).replace(/,/g,"")) || 0;   // Load report weights are in pounds

  const byId = new Map(), names = {}; let first = "", last = "";
  const cCall = head.indexOf("Actual call weight");
  for (const r of grid.slice(1)){
    if (r.length < head.length/2) continue;
    const ts = S(r[cDate]), d = ts.slice(0,10);
    if (d){ if(!first || d<first) first=d; if(d>last) last=d; }
    if (!S(r[cIng])) continue;                                   // pen-delivery row
    const rc = codeOf(r[cRecipe]), ing = codeOf(r[cIng]);
    names[rc] = nameOf(r[cRecipe]); if(!(ing in names)) names[ing] = nameOf(r[cIng]);
    let b = byId.get(S(r[cBatch]));
    if (!b){ b = {recipe:rc, time:ts, ing:{}}; byId.set(S(r[cBatch]), b); }
    if (ts && ts < b.time) b.time = ts;
    const x = b.ing[ing] = b.ing[ing] || {act:0, dry:0, call:0};
    x.act += lbs(r[cAct]); x.dry += lbs(r[cDry]); x.call += lbs(r[cCall]);
  }
  const batches = [...byId.values()].sort((a,b) => a.time < b.time ? -1 : a.time > b.time ? 1 : 0);
  const sum = (o, k) => Object.values(o).reduce((a,x) => a + (k ? x[k] : x), 0);

  const uses = {};
  batches.forEach(b => { uses[b.recipe] = uses[b.recipe] || new Set(); Object.keys(b.ing).forEach(i => uses[b.recipe].add(i)); });
  const recipes = Object.keys(uses);
  const premixRecipes = recipes.filter(rc => recipes.some(o => o !== rc && uses[o].has(rc)));
  const lactating = recipes.filter(rc => PREMIX_CODES.some(p => uses[rc].has(p)));
  const excluded = recipes.filter(rc => !lactating.includes(rc) && !premixRecipes.includes(rc));

  const monthStart = last.slice(0,8) + "01", start = first < monthStart ? monthStart : first;
  const made = code => batches.filter(b => b.recipe === code && sum(b.ing, "act") > 0);
  const notes = [], missing = {};
  // Each recipe's shares come only from its batches made in the counted month (pooled by weight).
  const pooled = code => {
    const all = made(code); let use = all.filter(b => b.time.slice(0,10) >= start);
    if (!use.length && all.length){
      use = [all[all.length-1]];
      notes.push(`No ${names[code] || code} batch was made in this period, so the last one before it (${usDate(use[0].time.slice(0,10))}) was used.`);
    }
    if (!use.length) return null;
    const tot = {act:0, dry:0, call:0}, ing = {};
    use.forEach(b => Object.entries(b.ing).forEach(([i, x]) => { const y = ing[i] = ing[i] || {act:0, dry:0, call:0};
      ["act","dry","call"].forEach(k => { y[k] += x[k]; tot[k] += x[k]; }); }));
    const frac = (i, k) => tot[k] > 0 && ing[i] ? ing[i][k] / tot[k] : 0;
    return { ing, frac, share: Object.fromEntries(Object.keys(ing).map(i => [i, frac(i, "dry")])) };
  };
  const rcp = {}, recipeOf = code => code in rcp ? rcp[code] : (rcp[code] = pooled(code));

  const before = {}, after = {}, card = {}, src = {};
  let herd = 0;
  const add = (o, k, v) => o[k] = (o[k] || 0) + v;
  const source = code => src[code] = src[code] || {quantityLb:0, bovarAF:0, bovaerAF:0, bovaerDM:0, intendedAF:0};

  // dm/af: lb of BOVAR reaching the cows. Dose follows the as-fed chain a verifier can redo from the report.
  const feedBovar = (dm, af, callAf, s) => {
    const r = recipeOf("BOVAR");
    if (!r){ add(missing, "BOVAR", dm); return; }
    card.BOVAR = card.BOVAR || {};
    Object.entries(r.share).forEach(([i, sh]) => { add(after, i, dm * sh); add(card.BOVAR, i, dm * sh); });
    const bov = r.ing.Bovaer || {act:0, dry:0}, bovaerAF = af * r.frac("Bovaer", "act");
    s.bovarAF += af; s.bovaerAF += bovaerAF; s.bovaerDM += bov.act > 0 ? bovaerAF * bov.dry / bov.act : 0;
    s.intendedAF += callAf * r.frac("Bovaer", "call");
  };
  const feedPremix = (p, x) => {
    const r = recipeOf(p);
    if (!r){ add(missing, p, x.dry); return; }
    card[p] = card[p] || {};
    Object.entries(r.share).forEach(([i, sh]) => { add(card[p], i, x.dry * sh); if (i !== "BOVAR") add(after, i, x.dry * sh); });
    const s = source(p); s.quantityLb += x.act;
    if (r.share.BOVAR) feedBovar(x.dry * r.share.BOVAR, x.act * r.frac("BOVAR", "act"), x.act * r.frac("BOVAR", "call"), s);
  };

  for (const b of batches){
    if (b.recipe === "BOVAR" || PREMIX_CODES.includes(b.recipe) || !lactating.includes(b.recipe)) continue;
    if (b.time.slice(0,10) < start) continue;
    for (const [i, x] of Object.entries(b.ing)){
      add(before, i, x.dry); herd += x.dry;
      if (PREMIX_CODES.includes(i)) feedPremix(i, x);
      else if (i === "BOVAR"){ const s = source("BOVAR"); s.quantityLb += x.act; feedBovar(x.dry, x.act, x.act, s); }
      else add(after, i, x.dry);
    }
  }

  if (start > first) notes.unshift(`Counted ${usDate(start)} – ${usDate(last)}. Earlier rows in the file are not used: each month is worked out from its own batches.`);

  const T = v => v / LB_PER_TON;
  const desc = code => names[code] || "";
  const list = o => Object.entries(o).filter(([,v]) => T(v) > 0.0005).sort((a,b) => b[1]-a[1]).map(([code, dm]) => ({code, desc:desc(code), dm:T(dm)}));
  const premixes = [...PREMIX_CODES, "BOVAR"].filter(p => card[p] && sum(card[p]) > 0).map(p => {
    const fed = sum(card[p]), mixedLb = made(p).filter(b => b.time.slice(0,10) >= start).reduce((a,b) => a + sum(b.ing, "dry"), 0);
    return { code:p, fedDM:T(fed), batchDM:T(mixedLb), rows: Object.entries(card[p]).sort((a,b) => b[1]-a[1])
      .map(([code, dm]) => ({code, desc:desc(code), pct: dm/fed*100, fedDM:T(dm)})) };
  });
  const missingList = Object.entries(missing).map(([code, dm]) => ({code, dm:T(dm)}));
  const sources = Object.values(src), bovaerAF = sum(sources.map(s => s.bovaerAF)), bovaerDM = sum(sources.map(s => s.bovaerDM));
  const intendedAF = sum(sources.map(s => s.intendedAF));
  const dose = bovaerAF > 0 && herd > 0 && !missingList.length
    ? { mgPerKg: bovaerAF*NOP_IN_BOVAER/herd*1e6, target: DOSE_TARGET, bovaerAsFed: T(bovaerAF), bovaerDmPct: bovaerDM/bovaerAF*100,
        basis: "actual", intended: intendedAF > 0 ? { mgPerKg: intendedAF*NOP_IN_BOVAER/herd*1e6, bovaerAsFed: T(intendedAF) } : null } : null;
  const calcSources = dose ? Object.entries(src).filter(([, s]) => s.bovaerAF > 0).map(([code, s]) => {
    const quantityLb = code === "BOVAR" ? s.bovarAF : s.quantityLb, inclusionPct = s.bovaerAF / s.bovarAF * 100;
    return { code, name: names[code] || code, quantityLb, dmPct: null, inclusionPct,
             nestedPct: code === "BOVAR" ? null : s.bovarAF / quantityLb * 100, bovaerLb: s.bovaerAF, nopMg: s.bovaerAF * NOP_IN_BOVAER * MG_PER_LB };
  }) : null;
  const dateRange = first ? `From ${usDate(start)} until ${usDate(last)}` : "";
  const beforeList = list(before), afterList = list(after);

  return {
    data: { dateRange, days: daysInRange(dateRange), totalHeads: 0, herdDM: T(herd), missing: missingList, premixes,
            reportTotal: 0, mixingDM: 0, dose, calculator: calcSources && calcSummary(calcSources, T(herd)),
            before: beforeList, after: afterList, beforeTotal: sum(beforeList.map(r => r.dm)), afterTotal: sum(afterList.map(r => r.dm)) },
    included: lactating.map(rc => names[rc] || rc), excluded: excluded.map(rc => names[rc] || rc), notes };
}

const toTons = (v,u) => { const n = parseFloat(v||0)||0; return (u && u.indexOf("100")>=0) ? n*100/2000 : n; };
function daysInRange(dr){
  const m = (dr||"").match(/(\d{1,2})\/(\d{1,2})\/(\d{4})/g);
  if (m && m.length>=2){ const p=s=>{const[a,b,c]=s.split("/").map(Number);return new Date(c,a-1,b);};
    return Math.round((p(m[1])-p(m[0]))/86400000)+1; }
  return null;
}

/* ---------- core calculation ---------- */
function build(recs){
  const dateRange = (recs.find(r=>r.DateRange)||{}).DateRange || (recs.find(r=>r._dr)||{})._dr || "";
  const days = daysInRange(dateRange);
  const animalPens = [...new Set(recs
    .filter(r=> (r.RowType==="Ingredient"||r.RowType==="SubTotal"))
    .map(r=>r.Pen).filter(p=> p && !NON_PEN.includes(p)))];

  const heads = {};
  recs.forEach(r=>{ if(animalPens.includes(r.Pen) && r.AvgHeads){ const h=parseFloat(r.AvgHeads); if(!isNaN(h)) heads[r.Pen]=h; }});
  const totalHeads = Object.values(heads).reduce((a,b)=>a+b,0);

  const desc = {};
  recs.forEach(r=>{ if(r.DisplayName && r.Description && !(r.DisplayName in desc)) desc[r.DisplayName]=r.Description; });

  const before = {};
  recs.forEach(r=>{ if(animalPens.includes(r.Pen) && r.RowType==="Ingredient")
      before[r.DisplayName] = (before[r.DisplayName]||0) + toTons(r.DryWt_Ton, r.Units); });

  const premixes = [], missing = [];
  // Scale a premix's batch recipe up to the DM actually fed; null if the report has no batches for it.
  const expand = (pmx, fedDM) => {
    const box = recs.filter(r=>r.Pen===pmx && r.RowType==="Ingredient");
    const batchDM = box.reduce((a,r)=>a+toTons(r.DryWt_Ton,r.Units),0);
    if(batchDM <= 0){ missing.push({code:pmx, dm:fedDM}); return null; }
    const factor = fedDM/batchDM;
    const rows = box.map(r=>{ const b=toTons(r.DryWt_Ton,r.Units);
      return { code:r.DisplayName, desc:desc[r.DisplayName]||"", pct:b/batchDM*100, fedDM:b*factor }; })
      .sort((a,b)=>b.fedDM-a.fedDM);
    const p = { code:pmx, fedDM, batchDM, factor, rows };
    premixes.push(p); return p;
  };

  const after = {};
  const add = (k,v)=> after[k]=(after[k]||0)+v;
  Object.entries(before).forEach(([k,v])=>{ if(!PREMIX_CODES.includes(k)) add(k,v); });
  PREMIX_CODES.forEach(pmx=>{ if((before[pmx]||0) > 0){ const p = expand(pmx, before[pmx]); if(p) p.rows.forEach(r=> add(r.code, r.fedDM)); }});
  // BOVAR sits inside the premixes, so it is broken out second, from its own batches.
  if((after.BOVAR||0) > 0){ const fed = after.BOVAR; delete after.BOVAR;
    const p = expand("BOVAR", fed); if(p) p.rows.forEach(r=> add(r.code, r.fedDM)); }

  Object.keys(after).forEach(k=>{ if(after[k]<=0.0005) delete after[k]; });
  Object.keys(before).forEach(k=>{ if(before[k]<=0.0005) delete before[k]; });

  const herdDM = recs.filter(r=>animalPens.includes(r.Pen) && r.RowType==="SubTotal")
                     .reduce((a,r)=>a+toTons(r.DryWt_Ton,r.Units),0);
  const reportTotal = recs.filter(r=>r.RowType==="GrandTotal").reduce((a,r)=>a+toTons(r.DryWt_Ton,r.Units),0);
  const mixingDM = recs.filter(r=>["PREMIX","2PRMIX","BOVAR"].includes(r.Pen) && r.RowType==="SubTotal")
                       .reduce((a,r)=>a+toTons(r.DryWt_Ton,r.Units),0);
  const list = obj => Object.entries(obj).sort((a,b)=>b[1]-a[1]).map(([code,dm])=>({code, desc:desc[code]||"", dm}));
  const sum  = obj => Object.values(obj).reduce((a,b)=>a+b,0);

  // 3-NOP is labelled on an as-fed basis, so convert Bovaer DM back using its DM% from the report.
  const bov = recs.filter(r=>r.DisplayName==="Bovaer" && r.RowType==="Ingredient");
  const bovDry = bov.reduce((a,r)=>a+toTons(r.DryWt_Ton,r.Units),0);
  const bovAct = bov.reduce((a,r)=>a+toTons(r.ActualWt_Ton,r.Units),0);
  const bovaerAsFed = after.Bovaer && bovDry > 0 && bovAct > 0 ? after.Bovaer * bovAct/bovDry : null;
  const dose = bovaerAsFed && herdDM > 0 && !missing.length
    ? { mgPerKg: bovaerAsFed*NOP_IN_BOVAER/herdDM*1e6, target: DOSE_TARGET, bovaerAsFed, bovaerDmPct: bovDry/bovAct*100 }
    : null;
  const calculator = dose ? calculatorInputs(recs, premixes, before, herdDM, bovAct, bovDry) : null;

  return { dateRange, days, totalHeads, herdDM, missing, premixes, reportTotal, mixingDM, dose, calculator,
           before:list(before), after:list(after), beforeTotal:sum(before), afterTotal:sum(after) };
}

/* ---------- Bovaer Dose Calculator inputs ---------- */
// The same numbers the Bovaer Dose Calculator takes in "Bovaer in Blend" mode, nested in a premix, all as-fed.
// Nested % is set so each source carries exactly the Bovaer the breakdown above found, so the calculator's dose matches.
function calculatorInputs(recs, premixes, before, herdDM, bovAct, bovDry){
  const penSum = (pen, k) => recs.filter(r => r.Pen === pen && r.RowType === "Ingredient")
                                 .reduce((a, r) => a + toTons(r[k], r.Units), 0);
  const bovarAF = penSum("BOVAR", "ActualWt_Ton"), bovarDM = penSum("BOVAR", "DryWt_Ton");
  const bovarBox = premixes.find(p => p.code === "BOVAR");
  const bovaerDmShare = ((bovarBox && bovarBox.rows.find(r => r.code === "Bovaer")) || {pct:0}).pct / 100;
  const inclusionPct = bovarAF > 0 ? bovAct / bovarAF * 100 : 0;
  const mgOf = bovaerAFtons => bovaerAFtons * LB_PER_TON * NOP_IN_BOVAER * MG_PER_LB;

  const sources = [];
  premixes.filter(p => PREMIX_CODES.includes(p.code)).forEach(p => {
    const bovarRow = p.rows.find(r => r.code === "BOVAR");
    if (!bovarRow || bovarRow.fedDM <= 0) return;
    const dmPct = penSum(p.code, "DryWt_Ton") / penSum(p.code, "ActualWt_Ton") * 100;
    const quantityLb = p.fedDM / (dmPct / 100) * LB_PER_TON;
    const bovaerAF = bovarRow.fedDM * bovaerDmShare * bovAct / bovDry;
    sources.push({ code: p.code, name: (recs.find(r => r.DisplayName === p.code && r.Description) || {}).Description || p.code,
      quantityLb, dmPct, inclusionPct, nestedPct: bovaerAF * LB_PER_TON / (quantityLb * inclusionPct / 100) * 100,
      bovaerLb: bovaerAF * LB_PER_TON, nopMg: mgOf(bovaerAF) });
  });
  if ((before.BOVAR || 0) > 0 && bovarDM > 0){            // Bovaer premix fed straight into a ration
    const bovaerAF = before.BOVAR * bovaerDmShare * bovAct / bovDry;
    sources.push({ code: "BOVAR", name: "Bovaer Premix", quantityLb: before.BOVAR / (bovarDM / bovarAF) * LB_PER_TON,
      dmPct: bovarDM / bovarAF * 100, inclusionPct, nestedPct: null, bovaerLb: bovaerAF * LB_PER_TON, nopMg: mgOf(bovaerAF) });
  }
  return calcSummary(sources, herdDM);
}
function calcSummary(sources, herdDMtons){
  const totalFeedLbDM = herdDMtons * LB_PER_TON, totalFeedKgDM = totalFeedLbDM * KG_PER_LB;
  const totalNopMg = sources.reduce((a, s) => a + s.nopMg, 0);
  return { sources, totalFeedLbDM, totalFeedKgDM, totalNopMg, mgPerKg: totalNopMg / totalFeedKgDM,
           labelMin: LABEL_MIN, labelMax: LABEL_MAX };
}

/* ---------- TapFeed Premixes / Recipes report (PDF, as text lines) ---------- */
const PMX_HEAD = /^(.*?)\s*(\S+)\s+([\d,]+\.\d+)\s+([\d.]+)%\s+([\d.]+)\s+([\d.]+)\s+([\d.]+)(?:\s*[\u2713\u2717])?$/;
const PMX_ING = /^(\d+)\s+(.*?)\s+(\S+)\s+([\d.]+)%\s+([\d.]+)\s+([\d.]+)%\s+([\d.]+)\s+([\d.]+)%$/;

// Weights print to 0.01 lb/animal and load % to 0.1%, so each share uses whichever of the two is less rounded.
function bestShare(part, whole, pct){
  const byWeight = part > 0 && whole > 0 && 0.005 / part <= 0.05 / Math.max(pct, 1e-9);
  return byWeight ? part / whole : pct / 100;
}

function parsePremixes(lines){
  const top = lines.slice(0, 6).map(l => l.trim());
  const kind = top.some(l => /^Premixes$/i.test(l)) ? "premixes" : top.some(l => /^Recipes$/i.test(l)) ? "recipes" : null;
  const generated = ((lines.find(l => /^Generated:/i.test(l.trim())) || "").replace(/^\s*Generated:\s*/i, "").match(/\d{1,2}\/\d{1,2}\/\d{4}/) || [""])[0];
  const raw = {}; let cur = null;
  for (const l of lines.map(s => s.trim())){
    const g = l.match(PMX_ING);
    if (g && cur){ cur.ings.push({ name:g[2], code:g[3], af:+g[5], afPct:+g[6], dm:+g[7], dmPct:+g[8] }); continue; }
    const h = l.match(PMX_HEAD);
    if (h && !/^Recipe\s+Display name/i.test(l)) cur = raw[h[2]] = { code:h[2], name:h[1].trim() || h[2], af:+h[5], dm:+h[6], ings:[] };
  }
  const recipes = {};
  Object.values(raw).forEach(r => {
    const by = {};
    r.ings.forEach(i => { const x = by[i.code] = by[i.code] || {name:i.name, af:0, afPct:0, dm:0, dmPct:0};
      x.af += i.af; x.afPct += i.afPct; x.dm += i.dm; x.dmPct += i.dmPct; });
    recipes[r.code] = { code:r.code, name:r.name,
      names: Object.fromEntries(Object.entries(by).map(([c, x]) => [c, x.name])),
      afShare: Object.fromEntries(Object.entries(by).map(([c, x]) => [c, bestShare(x.af, r.af, x.afPct)])),
      dmShare: Object.fromEntries(Object.entries(by).map(([c, x]) => [c, bestShare(x.dm, r.dm, x.dmPct)])) };
  });
  const bovarCode = (Object.values(recipes).find(r => "Bovaer" in r.afShare) || {}).code || null;
  const carriers = bovarCode ? Object.values(recipes).filter(r => bovarCode in r.afShare).map(r => r.code) : [];
  return { kind, generated, recipes, bovarCode, carriers };
}

/* ---------- TapFeed Ingredient usage report (CSV, per pen, names only, tons) ---------- */
// Columns are ingredient, pen, head count, actual t, dry t (the header row labels are shifted). No date range in the file.
// Bovaer intended = premix as fed x Bovaer premix % of the premix x Bovaer % of the Bovaer premix, all from the Premixes recipe.
function fromUsage(grid, pmx){
  const S = v => (v===undefined||v===null) ? "" : String(v).trim();
  const num = v => parseFloat(S(v).replace(/,/g, "")) || 0;
  const rows = grid.slice(1).filter(r => S(r[1]) && S(r[1]).toUpperCase() !== "TOTAL" && S(r[0]))
    .map(r => ({ ing:S(r[0]), pen:S(r[1]), heads:num(r[2]), act:num(r[3]), dry:num(r[4]) }));
  const byName = {}; Object.values(pmx.recipes).forEach(r => byName[r.name.toLowerCase()] = r);
  const carrierNames = pmx.carriers.map(c => pmx.recipes[c].name.toLowerCase());
  const pens = [...new Set(rows.map(r => r.pen))];
  const lactating = pens.filter(p => rows.some(r => r.pen === p && carrierNames.includes(r.ing.toLowerCase())));
  const excluded = pens.filter(p => !lactating.includes(p));

  const before = {}, after = {}, card = {}, src = {}, notes = [];
  const add = (o, k, v) => o[k] = (o[k] || 0) + v;
  const bovar = pmx.recipes[pmx.bovarCode];
  let herd = 0, heads = 0;
  const expand = (recipe, dm, into) => Object.entries(recipe.dmShare).forEach(([c, sh]) => {
    const name = recipe.names[c];
    if (c === pmx.bovarCode) Object.entries(bovar.dmShare).forEach(([c2, sh2]) => { add(into, bovar.names[c2], dm * sh * sh2); add(card[bovar.name] = card[bovar.name] || {}, bovar.names[c2], dm * sh * sh2); });
    else add(into, name, dm * sh);
  });
  lactating.forEach(p => {
    const pr = rows.filter(r => r.pen === p);
    heads += Math.max(...pr.map(r => r.heads));
    pr.forEach(r => {
      herd += r.dry; add(before, r.ing, r.dry);
      const recipe = byName[r.ing.toLowerCase()];
      if (recipe){
        card[recipe.name] = card[recipe.name] || {};
        Object.entries(recipe.dmShare).forEach(([c, sh]) => add(card[recipe.name], recipe.names[c], r.dry * sh));
        expand(recipe, r.dry, after);
        if (pmx.carriers.includes(recipe.code)){
          const s = src[recipe.code] = src[recipe.code] || { quantityLb:0 };
          s.quantityLb += r.act * LB_PER_TON;
        }
      } else add(after, r.ing, r.dry);
    });
  });

  const sum = o => Object.values(o).reduce((a, b) => a + b, 0);
  const list = o => Object.entries(o).filter(([, v]) => v > 0.0005).sort((a, b) => b[1] - a[1]).map(([code, dm]) => ({code, desc:"", dm}));
  const bovaerShare = bovar ? (bovar.afShare.Bovaer || 0) : 0;
  const sources = Object.entries(src).map(([code, s]) => {
    const nested = pmx.recipes[code].afShare[pmx.bovarCode], bovaerLb = s.quantityLb * nested * bovaerShare;
    return { code, name: pmx.recipes[code].name, quantityLb: s.quantityLb, dmPct: null, inclusionPct: bovaerShare * 100,
             nestedPct: nested * 100, bovaerLb, nopMg: bovaerLb * NOP_IN_BOVAER * MG_PER_LB };
  });
  const bovaerLb = sources.reduce((a, s) => a + s.bovaerLb, 0);
  const dose = bovaerLb > 0 && herd > 0
    ? { mgPerKg: bovaerLb / LB_PER_TON * NOP_IN_BOVAER / herd * 1e6, target: DOSE_TARGET, bovaerAsFed: bovaerLb / LB_PER_TON,
        bovaerDmPct: null, basis: "intended", intended: null } : null;
  notes.push("Ingredient usage reports don't show their date range. Check the export covered the full month.");
  notes.push(`Intended dose from the Premixes recipe${pmx.generated ? ` (printed ${pmx.generated})` : ""}: the recipe prints weights to 0.01 lb, so this is accurate to about ±1%. Make sure the recipe was the one in use for the whole period.`);
  const premixes = Object.entries(card).map(([name, o]) => { const fed = sum(o);
    return { code:name, fedDM:fed, batchDM:0, rows: Object.entries(o).sort((a, b) => b[1] - a[1]).map(([c, dm]) => ({code:c, desc:"", pct:dm/fed*100, fedDM:dm})) }; });
  const beforeList = list(before), afterList = list(after), total = l => l.reduce((a, r) => a + r.dm, 0);
  return {
    data: { dateRange:"", days:null, totalHeads:heads, herdDM:herd, missing:[], premixes, reportTotal:0, mixingDM:0, dose,
            calculator: dose ? calcSummary(sources, herd) : null, before:beforeList, after:afterList,
            beforeTotal: total(beforeList), afterTotal: total(afterList) },
    included: lactating, excluded, notes };
}

/* ---------- which file is which ---------- */
const KINDS = {
  load:      { label:"TapFeed Load report", role:"Everything needed: actual and recipe (intended) dose" },
  usage:     { label:"TapFeed Ingredient usage report", role:"What each pen was fed (needs the Premixes report for the dose)" },
  premixes:  { label:"TapFeed Premixes report", role:"What's in each premix (needs the Ingredient usage report)" },
  recipes:   { label:"TapFeed Recipes report", role:"Not needed: the Premixes report or Load report is" },
  summary:   { label:"TapFeed Ingredient/Premix Usage summary", role:"Can't be used: it double counts premix mixing and doesn't show what's in each premix" },
  tracker:   { label:"TMR Tracker ingredient report", role:"Older report: works if it includes the BOVAR batch section" },
  loadXlsx:  { label:"Load report in Excel format", role:"Re-export it as CSV" },
  unknown:   { label:"Not recognised", role:"Not a TapFeed or TMR Tracker report" },
};

// file: {name, content} where content is a string (CSV), ArrayBuffer/Uint8Array (Excel) or an array of text lines (PDF).
function identify(file, XLSX){
  const lower = file.name.toLowerCase();
  if (lower.endsWith(".pdf")){
    const p = parsePremixes(file.content || []);
    return { kind: p.kind || "unknown", premixes: p };
  }
  if (lower.endsWith(".csv")){
    const all = parseCSV(file.content), grid = all.filter(r => !String(r[0]||"").startsWith("#"));
    const head = (grid[0] || []).map(h => String(h).trim());
    const farmLine = all.find(r => /^#\s*Farm:/i.test(String(r[0]||"")));
    const farm = farmLine ? String(farmLine[0]).replace(/^#\s*Farm:\s*/i, "").trim() : "";
    if (head.includes("BatchId") && head.includes("Recipe") && head.includes("Dry weight")) return { kind:"load", farm };
    if (head[0] === "Ingredient" && head.includes("Loaded weight")) return { kind:"summary", farm };
    if (head.includes("Group") && head.includes("# Animals") && head.includes("Dry weight")) return { kind:"usage", farm, grid };
    if (head.includes("DryWt_Ton")) return { kind:"tracker" };
    return { kind:"unknown" };
  }
  if (lower.endsWith(".xls") || lower.endsWith(".xlsx")){
    const wb = XLSX.read(file.content, {type:"array"});
    const rows = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], {header:1, raw:true, defval:""});
    if (rows.some(r => r.some(c => String(c).trim() === "Load Report"))) return { kind:"loadXlsx" };
    return { kind: rowsToRecords(rows).length ? "tracker" : "unknown" };
  }
  return { kind:"unknown" };
}

// Picks the best calculation the files allow and says what's missing. Returns {files, result, basis, guidance}.
function analyze(files, XLSX){
  const found = files.map(f => Object.assign({ name:f.name }, identify(f, XLSX)));
  const pick = k => found.find(f => f.kind === k), src = f => files.find(x => x.name === f.name);
  const load = pick("load"), tracker = pick("tracker"), usage = pick("usage");
  const pmx = found.find(f => f.kind === "premixes" && f.premixes.carriers.length);
  const guidance = [];
  let result = null, basis = null, used = [];

  if (load){
    result = processReport(load.name, src(load).content, XLSX); basis = "actual"; used = [load.name];
    if (usage && pmx){
      const u = fromUsage(usage.grid, pmx.premixes);
      if (u.data.dose) result.info.crossCheck = { mgPerKg: u.data.dose.mgPerKg,
        label: "Ingredient usage × Premixes recipe (intended, as delivered). Only comparable if both reports cover the same dates." };
      used.push(usage.name, pmx.name);
    }
  } else if (tracker){
    result = processReport(tracker.name, src(tracker).content, XLSX); basis = "actual"; used = [tracker.name];
  } else if (usage && pmx){
    const u = fromUsage(usage.grid, pmx.premixes);
    if (!u.included.length) guidance.push(`None of the pens in the Ingredient usage report were fed ${pmx.premixes.carriers.map(c => pmx.premixes.recipes[c].name).join(" or ")}. Check both reports are from the same farm and month.`);
    else { result = { data:u.data, info:{ source:"Ingredient usage + Premixes report", farm:usage.farm, included:u.included, excluded:u.excluded, notes:u.notes } };
           basis = "intended"; used = [usage.name, pmx.name]; }
  }

  if (!result){
    if (usage && !pmx) guidance.push("Ingredient usage found. To work out the Bovaer dose, also add the Premixes report (PDF) for the same period, or use the Load report instead, which has everything in one file.");
    if (pick("premixes") && !usage) guidance.push("Premixes report found. Also add the Ingredient usage report (CSV) for the month, or use the Load report instead.");
    if (pick("premixes") && !pmx) guidance.push("The Premixes report doesn't list a Bovaer premix (no recipe contains Bovaer). Check it's the right farm.");
    if (!found.length || !guidance.length) guidance.push("Add the TapFeed Load report (CSV). It's the only file you need.");
  } else if (basis === "intended") {
    guidance.push("For the actual dose (what was loaded), add the TapFeed Load report.");
  }
  found.forEach(f => { f.label = KINDS[f.kind].label; f.role = KINDS[f.kind].role;
    f.status = used.includes(f.name) ? "used" : ["summary","loadXlsx","unknown"].includes(f.kind) ? "unusable"
      : !result && ["usage","premixes"].includes(f.kind) ? "needs another file" : "not needed";
    if (f.status === "used" && f.kind !== "load")
      f.role = load ? "Used as a cross-check (recipe dose)" : f.kind === "usage" ? "What each pen was fed" : "What's in each premix"; });
  return { files: found, result, basis, guidance };
}

/* ---------- entry point: file contents -> result ---------- */
// The older ingredient report codes the BOVAR blend as "Bovaer" everywhere except inside the BOVAR batch section.
const legacyBovar = recs => { recs.forEach(r => { if (r.DisplayName === "Bovaer" && r.Pen !== "BOVAR"){
  r.DisplayName = "BOVAR"; r.Description = "Bovaer Premix"; } }); return recs; };

// content: string for .csv, ArrayBuffer/Uint8Array for .xls/.xlsx.
function processReport(fileName, content, XLSX){
  const lower = fileName.toLowerCase();
  const isCsv = lower.endsWith(".csv"), isXls = lower.endsWith(".xls") || lower.endsWith(".xlsx");
  if (!isCsv && !isXls)
    throw new UserError("That isn't a feed report. Use the TapFeed Load report (CSV) or the TMR Tracker ingredient report (CSV or Excel).");

  if (isCsv){
    const all = parseCSV(content);
    const farmLine = all.find(r => /^#\s*Farm:/i.test(String(r[0]||"")));
    const farm = farmLine ? String(farmLine[0]).replace(/^#\s*Farm:\s*/i, "").trim() : "";
    const grid = all.filter(r => !String(r[0]||"").startsWith("#"));
    const head = (grid[0] || []).map(h => String(h).trim());
    if (head.includes("BatchId") && head.includes("Recipe") && head.includes("Dry weight")){
      const lr = fromLoadReport(grid);
      if (!lr.included.length)
        throw new UserError("No milking-cow rations were found in that Load report (no ration contains PREMIX or 2PRMIX). Check the date range and try again.");
      return { data: lr.data, info: { source:"TapFeed Load report", farm, included:lr.included, excluded:lr.excluded, notes:lr.notes } };
    }
    if (head[0] === "Ingredient" && head.includes("Loaded weight"))
      throw new UserError("That's the Ingredient/Premix Usage summary. It adds the premix mixing loads on top of what the cows ate and doesn't show what's in each premix, so it can't be used. Export the TapFeed Load report (CSV) instead.");
    const recs = toRecords(grid);
    if (!recs.length || !("DryWt_Ton" in recs[0]))
      throw new UserError("That CSV doesn't look like a TapFeed or TMR Tracker report. Export the TapFeed Load report (CSV) and try again.");
    return { data: build(legacyBovar(recs)), info: { source:"TMR Tracker ingredient report" } };
  }

  const wb = XLSX.read(content, {type:"array"});
  const rows = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], {header:1, raw:true, defval:""});
  if (rows.some(r => r.some(c => String(c).trim() === "Load Report")))
    throw new UserError("That's a Load report in Excel format. Export the TapFeed Load report as CSV instead.");
  let dateRange = "";
  for(const r of rows){ for(const c of r){ const s=String(c||""); if(/From .*until/i.test(s)){ dateRange=s.trim(); break; } } if(dateRange) break; }
  const recs = rowsToRecords(rows);
  if(!recs.length) throw new UserError("Couldn't find ingredient rows in that Excel file. Make sure it's the TMR Tracker ingredient report.");
  if(dateRange) recs.forEach(r=> r._dr = dateRange);
  return { data: build(legacyBovar(recs)), info: { source:"TMR Tracker ingredient report" } };
}

const uploadRows = data => data.after.map(r => [r.code, "", "", +r.dm.toFixed(3), "Dry Matter", "Lactating"]);
const outputName = fileName => "UpLook_Ingredient_Upload_" + fileName.replace(/\.(csv|xlsx?)$/i, "");

const api = { NON_PEN, PREMIX_CODES, HEADERS, NOP_IN_BOVAER, DOSE_TARGET, LABEL_MIN, LABEL_MAX, UserError, parseCSV, toRecords, rowsToRecords,
              fromLoadReport, toTons, daysInRange, build, calculatorInputs, processReport, uploadRows, outputName,
              parsePremixes, fromUsage, identify, analyze, KINDS };
if (typeof module !== "undefined" && module.exports) module.exports = api;
else root.TMRCore = api;
})(typeof globalThis !== "undefined" ? globalThis : this);
