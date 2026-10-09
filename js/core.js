/* Pure parsing + calculation. No DOM access, so it runs in the browser and under `node --test`. */
(function (root) {
"use strict";

const NON_PEN = ["PREMIX","2PRMIX","BOVAR","TOTAL",""];
const PREMIX_CODES = ["PREMIX","2PRMIX"];
const HEADERS = ["Producer Ingredient Name","Pounds/Head/Day","Pounds/Month","Tons/Month","Dry Matter/As Fed","Cow Type"];
const LB_PER_TON = 2000;
const NOP_IN_BOVAER = 0.10;   // Bovaer 10 label: 10% 3-NOP, as-fed basis
const DOSE_TARGET = 60;       // mg 3-NOP per kg DM intake

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

/* ---------- TMR Tracker Load report (CSV, one row per ingredient load or pen delivery) ---------- */
const codeOf = s => { const m = String(s).match(/\(([^()]*)\)\s*$/); return m ? m[1].trim() : String(s).trim(); };
const nameOf = s => String(s).replace(/\s*\([^()]*\)\s*$/, "").trim();
const usDate = iso => { const [y,m,d] = iso.split("-").map(Number); return `${m}/${d}/${y}`; };

// Each ration recipe becomes a "pen" so build() can treat both report types the same way.
// Lactating = rations that contain a Bovaer-carrying premix (PREMIX / 2PRMIX).
function loadReportToRecords(grid){
  const head = grid[0].map(h => String(h).trim());
  const cRecipe = head.indexOf("Recipe"), cIng = head.indexOf("Ingredient"),
        cDate = head.indexOf("Load/Delivery date and time"), cDry = head.indexOf("Dry weight"),
        cAct = head.indexOf("Actual weight");
  const S = v => (v===undefined||v===null) ? "" : String(v).trim();
  const lbs = v => parseFloat(S(v).replace(/,/g,"")) || 0;   // Load report weights are in pounds

  const dm = {}, act = {}, names = {}; let first = "", last = "";
  for (const r of grid.slice(1)){
    if (r.length < head.length/2) continue;
    const d = S(r[cDate]).slice(0,10);
    if (d){ if(!first || d<first) first=d; if(d>last) last=d; }
    if (!S(r[cIng])) continue;                                   // pen-delivery row
    const rc = codeOf(r[cRecipe]), ing = codeOf(r[cIng]);
    names[rc] = nameOf(r[cRecipe]); if(!(ing in names)) names[ing] = nameOf(r[cIng]);
    (dm[rc] = dm[rc] || {})[ing] = (dm[rc][ing] || 0) + lbs(r[cDry]);
    (act[rc] = act[rc] || {})[ing] = (act[rc][ing] || 0) + lbs(r[cAct]);
  }

  const recipes = Object.keys(dm);
  const premixRecipes = recipes.filter(rc => recipes.some(o => o !== rc && rc in dm[o]));
  const lactating = recipes.filter(rc => PREMIX_CODES.some(p => p in dm[rc]));
  const excluded = recipes.filter(rc => !lactating.includes(rc) && !premixRecipes.includes(rc));

  const recs = [];
  const total = o => Object.values(o).reduce((a,b)=>a+b,0);
  const push = (pen, type, code, dryLb, actLb) => recs.push({Pen:pen, AvgHeads:"", RowType:type, DisplayName:code,
      Description: code ? (names[code]||"") : "", ActualWt_Ton: actLb/LB_PER_TON, DryWt_Ton: dryLb/LB_PER_TON, Units:"Ton"});
  for (const rc of lactating){
    Object.keys(dm[rc]).forEach(ing => push(rc, "Ingredient", ing, dm[rc][ing], act[rc][ing]));
    push(rc, "SubTotal", "", total(dm[rc]), total(act[rc]));
  }
  for (const rc of [...PREMIX_CODES, "BOVAR"]) if (dm[rc])
    Object.keys(dm[rc]).forEach(ing => push(rc, "Ingredient", ing, dm[rc][ing], act[rc][ing]));
  if (recs.length && first) recs[0].DateRange = `From ${usDate(first)} until ${usDate(last)}`;

  return { recs,
    included: lactating.map(rc => names[rc] || rc),
    excluded: excluded.map(rc => names[rc] || rc) };
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

  return { dateRange, days, totalHeads, herdDM, missing, premixes, reportTotal, mixingDM, dose,
           before:list(before), after:list(after), beforeTotal:sum(before), afterTotal:sum(after) };
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
    throw new UserError("That isn't a TMR Tracker export. Use the CSV or Excel (.xls/.xlsx) file from the TMR Tracker.");

  if (isCsv){
    const grid = parseCSV(content).filter(r => !String(r[0]||"").startsWith("#"));
    const head = (grid[0] || []).map(h => String(h).trim());
    if (head.includes("BatchId") && head.includes("Recipe") && head.includes("Dry weight")){
      const lr = loadReportToRecords(grid);
      if (!lr.included.length)
        throw new UserError("No milking-cow rations were found in that Load report (no ration contains PREMIX or 2PRMIX). Check the date range and try again.");
      return { data: build(lr.recs), info: { source:"Load report", included:lr.included, excluded:lr.excluded } };
    }
    if (head[0] === "Ingredient" && head.includes("Loaded weight"))
      throw new UserError("That's the Ingredient/Premix Usage summary. It adds the premix mixing loads on top of what the cows ate and doesn't show what's in each premix, so it can't be used. Export the Load report (CSV) instead.");
    const recs = toRecords(grid);
    if (!recs.length || !("DryWt_Ton" in recs[0]))
      throw new UserError("That CSV doesn't look like a TMR Tracker report. Export the Load report (CSV) or the monthly ingredient report and try again.");
    return { data: build(legacyBovar(recs)), info: { source:"Ingredient report" } };
  }

  const wb = XLSX.read(content, {type:"array"});
  const rows = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], {header:1, raw:true, defval:""});
  if (rows.some(r => r.some(c => String(c).trim() === "Load Report")))
    throw new UserError("That's the Load report in Excel format. Export the Load report as CSV instead.");
  let dateRange = "";
  for(const r of rows){ for(const c of r){ const s=String(c||""); if(/From .*until/i.test(s)){ dateRange=s.trim(); break; } } if(dateRange) break; }
  const recs = rowsToRecords(rows);
  if(!recs.length) throw new UserError("Couldn't find ingredient rows in that Excel file. Make sure it's the TMR Tracker ingredient report.");
  if(dateRange) recs.forEach(r=> r._dr = dateRange);
  return { data: build(legacyBovar(recs)), info: { source:"Ingredient report" } };
}

const uploadRows = data => data.after.map(r => [r.code, "", "", +r.dm.toFixed(3), "Dry Matter", "Lactating"]);
const outputName = fileName => "UpLook_Ingredient_Upload_" + fileName.replace(/\.(csv|xlsx?)$/i, "");

const api = { NON_PEN, PREMIX_CODES, HEADERS, NOP_IN_BOVAER, DOSE_TARGET, UserError, parseCSV, toRecords, rowsToRecords,
              loadReportToRecords, toTons, daysInRange, build, processReport, uploadRows, outputName };
if (typeof module !== "undefined" && module.exports) module.exports = api;
else root.TMRCore = api;
})(typeof globalThis !== "undefined" ? globalThis : this);
