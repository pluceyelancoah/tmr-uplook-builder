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

// Batch by batch, in time order. Each BOVAR batch's recipe applies to premix mixed until the next BOVAR batch,
// and each premix batch is fed out first-in, first-out. If the file starts before the last month in it, only that
// month is counted; earlier rows are batch history (e.g. the BOVAR batch in use on the 1st).
// Lactating = rations that contain a Bovaer-carrying premix (PREMIX / 2PRMIX).
function fromLoadReport(grid){
  const head = grid[0].map(h => String(h).trim());
  const cBatch = head.indexOf("BatchId"), cRecipe = head.indexOf("Recipe"), cIng = head.indexOf("Ingredient"),
        cDate = head.indexOf("Load/Delivery date and time"), cDry = head.indexOf("Dry weight"),
        cAct = head.indexOf("Actual weight");
  const S = v => (v===undefined||v===null) ? "" : String(v).trim();
  const lbs = v => parseFloat(S(v).replace(/,/g,"")) || 0;   // Load report weights are in pounds

  const byId = new Map(), names = {}; let first = "", last = "";
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
    const x = b.ing[ing] = b.ing[ing] || {act:0, dry:0};
    x.act += lbs(r[cAct]); x.dry += lbs(r[cDry]);
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
  const made = code => batches.filter(b => b.recipe === code && sum(b.ing, "dry") > 0);
  const bovarRecipe = b => { const dm = sum(b.ing, "dry"), bov = b.ing.Bovaer || {act:0, dry:0};
    return { time:b.time, dmFrac: dm / sum(b.ing, "act"), bovaerAfPerDm: bov.dry > 0 ? bov.act / bov.dry : 0,
             share: Object.fromEntries(Object.entries(b.ing).map(([i,x]) => [i, x.dry / dm])) }; };
  const bovarMade = made("BOVAR"), firstBovar = bovarMade.length ? bovarRecipe(bovarMade[0]) : null;
  const premixMade = Object.fromEntries(PREMIX_CODES.map(p => [p, made(p)]));

  const before = {}, after = {}, card = {BOVAR:{}}, missing = {}, early = {}, notes = [];
  const src = {}; let herd = 0, bovar = null, borrowed = 0;
  const add = (o, k, v) => o[k] = (o[k] || 0) + v;
  const source = code => src[code] = src[code] || {quantityLb:0, bovarAF:0, bovaerAF:0, bovaerDM:0};

  // BOVAR (lb DM) broken into Bovaer and its carrier with the BOVAR recipe that was in use.
  const feedBovar = (dm, recipe, s) => {
    if (!recipe){ add(missing, "BOVAR", dm); return; }
    Object.entries(recipe.share).forEach(([i, sh]) => { add(after, i, dm * sh); add(card.BOVAR, i, dm * sh); });
    s.bovarAF += dm / recipe.dmFrac;
    s.bovaerDM += dm * (recipe.share.Bovaer || 0);
    s.bovaerAF += dm * (recipe.share.Bovaer || 0) * recipe.bovaerAfPerDm;
  };
  const queue = Object.fromEntries(PREMIX_CODES.map(p => [p, []])), lastMix = {};
  const feedPremix = (p, dm, counted) => {
    if (!premixMade[p].length){ if (counted) add(missing, p, dm); return; }
    while (dm > 1e-9){
      let mix = queue[p][0], take = dm;
      if (mix) take = Math.min(dm, mix.left);
      else { mix = lastMix[p] || premixMade[p].map(premixRecipe)[0]; if (counted) add(early, p, dm); }  // fed before any batch on hand
      if (counted){
        card[p] = card[p] || {};
        Object.entries(mix.share).forEach(([i, sh]) => { add(card[p], i, take * sh); if (i !== "BOVAR") add(after, i, take * sh); });
        if (mix.share.BOVAR) feedBovar(take * mix.share.BOVAR, mix.bovar, source(p));
      }
      if (queue[p][0]){ mix.left -= take; if (mix.left <= 1e-9) queue[p].shift(); }
      dm -= take;
    }
  };
  const premixRecipe = b => { const dm = sum(b.ing, "dry");
    return { left: dm, share: Object.fromEntries(Object.entries(b.ing).map(([i,x]) => [i, x.dry / dm])),
             bovar: b.ing.BOVAR ? (bovar || firstBovar) : null }; };

  for (const b of batches){
    const counted = b.time.slice(0,10) >= start;
    if (b.recipe === "BOVAR"){ if (sum(b.ing, "dry") > 0) bovar = bovarRecipe(b); continue; }
    if (PREMIX_CODES.includes(b.recipe)){
      if (sum(b.ing, "dry") <= 0) continue;
      if (b.ing.BOVAR && !bovar && firstBovar) borrowed++;
      const mix = premixRecipe(b); queue[b.recipe].push(mix); lastMix[b.recipe] = mix;
      continue;
    }
    if (!lactating.includes(b.recipe)) continue;
    for (const [i, x] of Object.entries(b.ing)){
      if (counted){ add(before, i, x.dry); herd += x.dry; }
      if (PREMIX_CODES.includes(i)){ if (counted) source(i).quantityLb += x.act; feedPremix(i, x.dry, counted); }
      else if (i === "BOVAR"){ if (counted) feedBovar(x.dry, bovar || firstBovar, source("BOVAR")); }
      else if (counted) add(after, i, x.dry);
    }
  }

  if (start > first) notes.push(`Counted ${usDate(start)} – ${usDate(last)}. Rows from ${usDate(first)} were used only as batch history (premix and Bovaer premix carried into the month).`);
  if (borrowed && firstBovar) notes.push(`${borrowed} premix batch${borrowed > 1 ? "es were" : " was"} mixed before the first Bovaer premix batch in this file (${usDate(firstBovar.time.slice(0,10))}) and given that batch's recipe. For exact numbers, export from the last Bovaer premix batch before the month.`);
  Object.entries(early).forEach(([p, dm]) => { const fed = before[p] || 0;
    if (fed && dm / fed > 0.01) notes.push(`${(dm/LB_PER_TON).toFixed(1)} t of ${names[p] || p} was fed before its first batch in this file and given the nearest batch's recipe.`); });

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
  const dose = bovaerAF > 0 && herd > 0 && !missingList.length
    ? { mgPerKg: bovaerAF*NOP_IN_BOVAER/herd*1e6, target: DOSE_TARGET, bovaerAsFed: T(bovaerAF), bovaerDmPct: bovaerDM/bovaerAF*100 } : null;
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
              fromLoadReport, toTons, daysInRange, build, calculatorInputs, processReport, uploadRows, outputName };
if (typeof module !== "undefined" && module.exports) module.exports = api;
else root.TMRCore = api;
})(typeof globalThis !== "undefined" ? globalThis : this);
