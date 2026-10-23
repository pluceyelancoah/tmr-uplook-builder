/* Page wiring: reads the dropped file, renders the summary and working, builds the UpLook download. */
(function () {
"use strict";
const { HEADERS, UserError, analyze, uploadRows, outputName } = window.TMRCore;

let lastRows = null, lastName = "uplook_ingredient_usage", lastData = null, lastInfo = null;
const files = [];   // {name, content}: text for CSV, ArrayBuffer for Excel, text lines for PDF

const $ = id => document.getElementById(id);
const esc = v => String(v ?? "").replace(/[&<>"']/g, c => ({"&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;","'":"&#39;"}[c]));
const fmt = (n,d)=> Number(n).toLocaleString(undefined,{minimumFractionDigits:d,maximumFractionDigits:d});

// "From 9/1/2026 until 9/30/2026" -> "Sep 1–30, 2026"
function period(dateRange, days){
  const m = (dateRange||"").match(/(\d{1,2})\/(\d{1,2})\/(\d{4})/g);
  if (!m || m.length < 2) return "Period not found in the report";
  const [a, b] = m.slice(0,2).map(s => { const [mo,d,y] = s.split("/").map(Number); return new Date(y, mo-1, d); });
  const mon = d => d.toLocaleString("en-US", {month:"short"});
  const text = a.getFullYear() !== b.getFullYear()
    ? `${mon(a)} ${a.getDate()}, ${a.getFullYear()} – ${mon(b)} ${b.getDate()}, ${b.getFullYear()}`
    : a.getMonth() !== b.getMonth()
      ? `${mon(a)} ${a.getDate()} – ${mon(b)} ${b.getDate()}, ${b.getFullYear()}`
      : `${mon(a)} ${a.getDate()}–${b.getDate()}, ${b.getFullYear()}`;
  return days ? `${text} (${days} days)` : text;
}

const BTN_OK = "bg-primary text-primary-foreground";
const BTN_WARN = "bg-amber-600 text-white";

function show(data, info, fileName){
  $("result").classList.remove("hidden");
  $("error").classList.add("hidden");
  $("r-file").textContent = fileName;
  $("r-period").textContent = `${info.source} · ${data.dateRange ? period(data.dateRange, data.days) : "period not shown in the report"}`;
  $("after-count").textContent = `· ${data.after.length} feeds`;

  const d = data.dose;
  $("source").innerHTML = (info.included
    ? `<p><b>Milking-cow ${data.dateRange ? "rations" : "pens"} included</b> (fed a premix that carries Bovaer): ${info.included.map(esc).join(", ")}.</p>`
      + (info.excluded.length ? `<p><b>Not included:</b> ${info.excluded.map(esc).join(", ")}.</p>` : "")
    : "")
    + (d ? `<p><b>Dose:</b> ${fmt(d.bovaerAsFed*2000,1)} lb Bovaer as fed`
      + (d.bovaerDmPct ? ` (${fmt(d.bovaerDmPct,0)}% DM in the report)` : ` (from the Premixes recipe)`) + ` × 10% 3-NOP `
      + `÷ ${fmt(data.herdDM*2000,0)} lb dry matter = ${fmt(d.mgPerKg,1)} mg/kg.</p>` : "")
    + (d && d.intended ? `<p><b>Recipe (intended) dose:</b> the same feeding with each batch's recipe call instead of what was loaded: `
      + `${fmt(d.intended.bovaerAsFed*2000,1)} lb Bovaer = ${fmt(d.intended.mgPerKg,1)} mg/kg.</p>` : "")
    + (info.crossCheck ? `<p><b>Cross-check:</b> ${esc(info.crossCheck.label)} ${fmt(info.crossCheck.mgPerKg,1)} mg/kg.</p>` : "");

  const rd = $("r-dose"), rds = $("r-dose-sub");
  $("r-dose-label").textContent = d && d.basis === "intended" ? "Bovaer dose (recipe, intended)" : "Bovaer dose";
  if (d){
    const ok = d.mgPerKg >= d.target;
    rd.textContent = fmt(d.mgPerKg,1);
    rd.className = "text-3xl font-semibold tnum " + (ok ? "text-success" : "text-danger");
    rds.textContent = `${ok ? "At or above" : "Below"} target of ${d.target} · ${fmt(d.mgPerKg/d.target*100,0)}% of target`
      + (d.intended ? ` · recipe (intended): ${fmt(d.intended.mgPerKg,1)}` : "");
  } else {
    rd.textContent = "—"; rd.className = "text-3xl font-semibold tnum";
    rds.textContent = data.missing.length ? "Can't calculate: premix batches missing" : "No Bovaer in this report";
  }

  const incomplete = data.missing.length > 0;
  const btn = $("dl-xlsx");
  btn.classList.remove(...BTN_OK.split(" "), ...BTN_WARN.split(" "));
  btn.classList.add(...(incomplete ? BTN_WARN : BTN_OK).split(" "));
  $("dl-label").textContent = incomplete ? "Download anyway — incomplete (.xlsx)" : "Download for UpLook (.xlsx)";

  const miss = $("missing");
  if (data.missing.length){
    const l = data.missing.map(m=>`${esc(m.code)} (${m.dm.toFixed(1)} t)`).join(", ");
    miss.className = "alert-warn"; miss.classList.remove("hidden");
    miss.innerHTML = `⚠ This report is missing a premix recipe: ${l}.`
      + `<span class="sub">That feed can't be broken into ingredients, so it is NOT in the upload and the dose can't be calculated. `
      + `Re-export the report so it includes that premix's batches before uploading to UpLook.</span>`;
  } else miss.classList.add("hidden");

  const notes = $("notes"), nl = info.notes || [];
  notes.innerHTML = nl.map(t => `<p>${esc(t)}</p>`).join("");
  notes.classList.toggle("hidden", !nl.length);

  // premix breakdowns
  const pm = $("premixes"); pm.innerHTML="";
  if(!data.premixes.length) pm.innerHTML = '<p class="text-sm text-muted-foreground">No premixes were fed to the herd this period.</p>';
  data.premixes.forEach(p=>{
    const card = document.createElement("div");
    card.className = "rounded-lg border border-border bg-muted/20 p-4";
    const rows = p.rows.map(r=>`<tr class="border-t border-border/70">
        <td class="py-1.5 pr-3 font-medium text-brand tnum">${esc(r.code)}</td>
        <td class="py-1.5 pr-3 text-muted-foreground">${esc(r.desc)}</td>
        <td class="py-1.5 pr-3 text-right tnum">${fmt(r.pct,1)}%</td>
        <td class="py-1.5 text-right tnum">${fmt(r.fedDM,3)}</td></tr>`).join("");
    card.innerHTML = `<div class="flex flex-wrap items-baseline justify-between gap-x-4">
        <h4 class="font-semibold text-brand">${esc(p.code)}</h4>
        <span class="text-[13px] text-muted-foreground">${fmt(p.fedDM,2)} t DM fed${p.batchDM ? ` · ${fmt(p.batchDM,2)} t DM mixed this period` : ""}</span></div>
      <div class="overflow-x-auto"><table class="w-full text-sm mt-2"><thead><tr class="text-muted-foreground text-xs uppercase tracking-wide">
        <th class="text-left font-medium pb-1">Ingredient</th><th class="text-left font-medium pb-1">Report name</th>
        <th class="text-right font-medium pb-1">% of mix</th><th class="text-right font-medium pb-1">Tons fed (DM)</th>
      </tr></thead><tbody>${rows}</tbody></table></div>`;
    pm.appendChild(card);
  });

  renderTable("before", data.before, data.beforeTotal);
  renderTable("after",  data.after,  data.afterTotal);

  const rn = $("reportnote");
  if(data.reportTotal > data.beforeTotal*1.001){
    rn.classList.remove("hidden");
    rn.innerHTML = `<b>Why this is lower than the report's printed total:</b> the report shows `
      + `${fmt(data.reportTotal,1)} t, but that also counts the premix <i>mixing</i> batches `
      + `(${fmt(data.mixingDM,1)} t — the same corn, canola, etc. logged again when the premixes were made). `
      + `Counting each feed once gives ${fmt(data.beforeTotal,1)} t — what the cows actually ate.`;
  } else rn.classList.add("hidden");

  const gap = data.herdDM ? Math.abs(data.afterTotal-data.herdDM)/data.herdDM*100 : 0;
  const recon = $("recon");
  if(!data.herdDM){ recon.className="alert-warn"; recon.innerHTML="⚠ No whole-herd totals were found in this file to check against."; }
  else if(gap < 0.3 && !data.missing.length){ recon.className="alert-ok";
    recon.innerHTML = `✓ Totals match: the upload has ${fmt(data.afterTotal,1)} t DM and the cows ate ${fmt(data.herdDM,1)} t DM. `
      + `Premixes were broken out without changing the total.`; }
  else { recon.className="alert-warn";
    recon.innerHTML = `⚠ Please double-check: the upload has ${fmt(data.afterTotal,1)} t DM but the cows ate ${fmt(data.herdDM,1)} t DM.`
      + (data.missing.length ? `<span class="sub">This is from the missing premix recipe below.</span>` : ``); }

  lastRows = uploadRows(data);
  lastData = data; lastInfo = info;
  $("dl-pdf").classList.toggle("hidden", !data.calculator);
}

function renderTable(id, rows, total){
  const t = $(id);
  t.innerHTML = `<thead class="sticky top-0"><tr class="bg-muted text-muted-foreground text-xs uppercase tracking-wide">
      <th class="text-left font-medium px-3 py-2">Ingredient</th>
      <th class="text-left font-medium px-3 py-2">Report name</th>
      <th class="text-right font-medium px-3 py-2">Tons / Month (DM)</th></tr></thead>`;
  const tb = document.createElement("tbody");
  rows.forEach((r,i)=>{ const tr=document.createElement("tr");
    tr.className = "border-t border-border " + (i%2? "bg-muted/20":"");
    tr.innerHTML = `<td class="px-3 py-1.5 font-medium text-brand tnum">${esc(r.code)}</td>`
      + `<td class="px-3 py-1.5 text-muted-foreground">${esc(r.desc)}</td>`
      + `<td class="px-3 py-1.5 text-right tnum">${fmt(r.dm,3)}</td>`;
    tb.appendChild(tr); });
  t.appendChild(tb);
  const tf = document.createElement("tfoot");
  tf.innerHTML = `<tr class="border-t-2 border-brand/30 bg-muted/50 font-semibold">
      <td class="px-3 py-2">TOTAL</td><td></td><td class="px-3 py-2 text-right tnum">${fmt(total,3)}</td></tr>`;
  t.appendChild(tf);
}

function fail(msg, detail){
  $("error").classList.remove("hidden");
  $("result").classList.add("hidden");
  $("error-msg").textContent = msg;
  $("error-detail").textContent = detail ? ("If this keeps happening, send this to Elanco: " + detail) : "";
}

/* ---------- PDF text (pdf.js is only loaded when a PDF is added) ---------- */
let pdfjs = null;
const loadScript = src => new Promise((ok, bad) => { const s = document.createElement("script"); s.src = src;
  s.onload = ok; s.onerror = () => bad(new Error("Couldn't load " + src)); document.head.appendChild(s); });
// The worker script loaded first runs pdf.js on the page itself, which also works when the page is opened as a file.
const pdfLib = () => pdfjs = pdfjs || loadScript("vendor/pdf.worker.min.js").then(() => loadScript("vendor/pdf.min.js")).then(() => window.pdfjsLib);

async function pdfLines(buffer){
  const lib = await pdfLib(), doc = await lib.getDocument({ data: new Uint8Array(buffer) }).promise, lines = [];
  for (let p = 1; p <= doc.numPages; p++){
    const items = (await (await doc.getPage(p)).getTextContent()).items.filter(i => i.str.trim()), rows = new Map();
    items.forEach(i => { const y = Math.round(i.transform[5]); (rows.get(y) || rows.set(y, []).get(y)).push(i); });
    [...rows.entries()].sort((a, b) => b[0] - a[0])
      .forEach(([, r]) => lines.push(r.sort((a, b) => a.transform[4] - b.transform[4]).map(i => i.str.trim()).join(" ")));
  }
  return lines;
}

const readFile = file => new Promise((ok, bad) => {
  const rd = new FileReader(), lower = file.name.toLowerCase();
  rd.onerror = () => bad(rd.error);
  rd.onload = () => lower.endsWith(".pdf") ? pdfLines(rd.result).then(ok, bad) : ok(rd.result);
  if (lower.endsWith(".csv")) rd.readAsText(file); else rd.readAsArrayBuffer(file);
});

const STATUS = { "used": ["✓", "text-success"], "needs another file": ["…", "text-amber-600"],
                 "not needed": ["–", "text-muted-foreground"], "unusable": ["✕", "text-danger"] };

function renderFiles(a){
  const box = $("files");
  box.classList.toggle("hidden", !files.length);
  $("file-list").innerHTML = a.files.map(f => { const [icon, cls] = STATUS[f.status];
    return `<li class="flex gap-3 items-baseline"><span class="${cls} font-semibold w-4">${icon}</span>`
      + `<span class="min-w-0"><span class="font-medium">${esc(f.label)}</span> <span class="text-muted-foreground break-all">· ${esc(f.name)}</span>`
      + `<span class="block text-[13px] text-muted-foreground">${esc(f.status === "not needed" && f.kind !== "recipes" ? "Not needed: the Load report already has this" : f.role)}</span></span></li>`; }).join("");
  $("guidance").innerHTML = a.guidance.map(g => `<p>→ ${esc(g)}</p>`).join("");
  $("guidance").classList.toggle("hidden", !a.guidance.length);
}

function run(){
  try{
    const a = analyze(files, window.XLSX);
    renderFiles(a);
    if (!a.result){ $("result").classList.add("hidden"); $("error").classList.add("hidden"); return; }
    const main = a.files.find(f => f.status === "used");
    lastName = outputName(main.name);
    show(a.result.data, a.result.info, a.files.filter(f => f.status === "used").map(f => f.name).join(" + "));
  }catch(err){
    if (err instanceof UserError) fail(err.message);
    else fail("Something went wrong reading those files. Please re-export them and try again.", err.message);
  }
}

async function handleFiles(list){
  $("drop-title").textContent = "Reading…";
  for (const file of list){
    try{
      const content = await readFile(file), i = files.findIndex(f => f.name === file.name);
      if (i >= 0) files[i] = { name:file.name, content }; else files.push({ name:file.name, content });
    }catch(err){ fail(`Couldn't read ${file.name}. Please re-export it and try again.`, err.message); }
  }
  $("drop-title").textContent = files.length > 1 ? `${files.length} files added` : (files[0] || {}).name || "Drop your TapFeed Load report here";
  $("drop-sub").textContent = "Drop or click to add more files";
  run();
}

$("clear-files").onclick = () => { files.length = 0; lastRows = lastData = null;
  $("files").classList.add("hidden"); $("result").classList.add("hidden"); $("error").classList.add("hidden");
  $("drop-title").textContent = "Drop your TapFeed Load report here"; $("drop-sub").textContent = "or click to choose it"; };

const drop = $("drop"), fileInput = $("file");
drop.onclick = ()=> fileInput.click();
drop.addEventListener("keydown", e=>{ if(e.key==="Enter"||e.key===" "){ e.preventDefault(); fileInput.click(); }});
fileInput.onchange = e => { if(e.target.files.length) handleFiles([...e.target.files]); fileInput.value = ""; };
["dragover","dragenter"].forEach(ev=> drop.addEventListener(ev, e=>{e.preventDefault(); drop.classList.add("border-brand","bg-brand/5");}));
["dragleave","drop"].forEach(ev=> drop.addEventListener(ev, e=>{e.preventDefault(); drop.classList.remove("border-brand","bg-brand/5");}));
drop.addEventListener("drop", e=>{ if(e.dataTransfer.files.length) handleFiles([...e.dataTransfer.files]); });

const downloadPdf = () => { if (lastData) window.TMRPdf.download(lastData, lastInfo, lastName); };

$("dl-xlsx").onclick = ()=>{
  if(!lastRows) return;
  const ws = XLSX.utils.aoa_to_sheet([HEADERS, ...lastRows]);
  ws["!cols"] = [{wch:24},{wch:15},{wch:15},{wch:14},{wch:17},{wch:13}];
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "Sheet1");
  XLSX.writeFile(wb, lastName+".xlsx");
  downloadPdf();
};
$("dl-csv").onclick = ()=>{
  if(!lastRows) return;
  const cell = v => /[",\n]/.test(v) ? '"'+String(v).replace(/"/g,'""')+'"' : v;
  const csv = [HEADERS, ...lastRows].map(r=>r.map(cell).join(",")).join("\r\n");
  const blob = new Blob([csv], {type:"text/csv"});
  const a = document.createElement("a"); a.href = URL.createObjectURL(blob); a.download = lastName+".csv"; a.click();
  downloadPdf();
};
$("dl-pdf").onclick = downloadPdf;
})();
