const test = require("node:test");
const assert = require("node:assert/strict");
const XLSX = require("xlsx");
const C = require("../js/core.js");

const close = (a, b, msg) => assert.ok(Math.abs(a - b) < 1e-9, `${msg}: ${a} != ${b}`);
const byCode = data => Object.fromEntries(data.after.map(r => [r.code, r.dm]));

// Weights in lb. BOVAR = 9% Bovaer DM (Bovaer itself 90% DM); PREMIX = 40% BOVAR; HIGH COW eats 500 lb PREMIX of a 1,000 lb batch.
const LOAD_REPORT = [
  "# Farm: TEST DAIRY",
  "# Generated: 2026-10-06 13:39:20",
  "BatchId,Recipe,BaseBatchId,Ingredient,Pen,Load/Delivery date and time,Actual call weight,Dry call weight,Actual weight,Dry weight,Error amount,Error %,Planned cost,Cost,Cost error,Operator",
  "b1,Bovaer Premix (BOVAR),,GROUND CORN (GRDCRN),,2026-09-01 08:00:00,0,0,1000,910,0,0,0,0,0,A",
  "b1,Bovaer Premix (BOVAR),,Bovaer (Bovaer),,2026-09-01 08:05:00,0,0,100,90,0,0,0,0,0,A",
  "b1,Bovaer Premix (BOVAR),,,Bovaer Premix (BOVAR),2026-09-01 08:10:00,0,,1100,,0,0,0,0,0,A",
  "p1,Premix MILK COWS (PREMIX),,CANOLA (CANOLA),,2026-09-02 08:00:00,0,0,700,600,0,0,0,0,0,A",
  "p1,Premix MILK COWS (PREMIX),,Bovaer Premix (BOVAR),,2026-09-02 08:05:00,0,0,450,400,0,0,0,0,0,A",
  "p1,Premix MILK COWS (PREMIX),,,Premix MILK COWS (PREMIX),2026-09-02 08:10:00,0,,1150,,0,0,0,0,0,A",
  "h1,HIGH COW (HIGCOW),,Premix MILK COWS (PREMIX),,2026-09-03 08:00:00,0,0,570,500,0,0,0,0,0,A",
  "h1,HIGH COW (HIGCOW),,Corn silage 2025 (CORSL),,2026-09-03 08:05:00,0,0,4500,1500,0,0,0,0,0,A",
  "h1,HIGH COW (HIGCOW),,Water (H2O),,2026-09-03 08:06:00,0,0,300,0.00,0,0,0,0,0,A",
  "h1,HIGH COW (HIGCOW),,,4 - High Holstein (4),2026-09-03 08:10:00,0,,5370,,0,0,0,0,0,A",
  "d1,DRY COW (DRYCOW),,STRAW (STRAW),,2026-09-30 08:00:00,0,0,1100,1000,0,0,0,0,0,A",
  "d1,DRY COW (DRYCOW),,,20 - Dry Cows (20),2026-09-30 08:10:00,0,,1100,,0,0,0,0,0,A",
].join("\r\n");

test("Load report: premixes and BOVAR broken out, totals reconcile, dose on as-fed Bovaer", () => {
  const { data, info } = C.processReport("Load_report.csv", LOAD_REPORT, XLSX);
  assert.equal(info.source, "TapFeed Load report");
  assert.deepEqual(info.included, ["HIGH COW"]);
  assert.deepEqual(info.excluded, ["DRY COW"]);
  assert.deepEqual(data.premixes.map(p => p.code), ["PREMIX", "BOVAR"]);

  assert.equal(data.dateRange, "From 9/1/2026 until 9/30/2026");
  assert.equal(data.days, 30);
  close(data.herdDM, 1, "herd DM t");
  close(data.afterTotal, data.herdDM, "after = herd");

  const t = byCode(data);
  assert.deepEqual(Object.keys(t).sort(), ["Bovaer", "CANOLA", "CORSL", "GRDCRN"]);
  close(t.CORSL, 0.75, "CORSL");
  close(t.CANOLA, 0.15, "CANOLA");
  close(t.Bovaer, 0.009, "Bovaer DM");
  close(t.GRDCRN, 0.091, "GRDCRN");

  // 0.009 t DM / 90% DM = 0.01 t as fed; x 10% 3-NOP / 1 t DM eaten = 1,000 mg/kg
  close(data.dose.bovaerAsFed, 0.01, "Bovaer as fed");
  close(data.dose.mgPerKg, 1000, "dose");
});

test("Load report: Bovaer Dose Calculator inputs reproduce the dose", () => {
  const { data, info } = C.processReport("Load_report.csv", LOAD_REPORT, XLSX);
  assert.equal(info.farm, "TEST DAIRY");
  const calc = data.calculator, [s] = calc.sources;
  assert.equal(calc.sources.length, 1);
  assert.equal(s.code, "PREMIX");
  close(s.inclusionPct, 100 / 1100 * 100, "Bovaer % of BOVAR, as fed");
  close(s.quantityLb, 570, "PREMIX loaded into rations, as fed lb");
  close(s.bovaerLb, 20, "Bovaer lb");
  close(calc.totalFeedLbDM, 2000, "total feed lb DM");
  // Calculator math: lb x 0.4536 kg x (100,000 mg/kg x inclusion% x nested%) / kg DM
  const mg = s.quantityLb * 0.45359237 * 100000 * s.inclusionPct / 100 * s.nestedPct / 100;
  assert.ok(Math.abs(mg / calc.totalFeedKgDM - data.dose.mgPerKg) < 1e-6, "calculator dose = parser dose");
  close(calc.mgPerKg, data.dose.mgPerKg, "calc mg/kg");
});

// Export started early (Aug 20) to pick up the BOVAR batch in use on Sep 1. BOVAR A = 10% Bovaer DM, B = 20%.
const HEAD = "BatchId,Recipe,BaseBatchId,Ingredient,Pen,Load/Delivery date and time,Actual call weight,Dry call weight,Actual weight,Dry weight,Error amount,Error %,Planned cost,Cost,Cost error,Operator";
const row = (id, recipe, ing, ts, act, dry) => `${id},${recipe},,${ing},,${ts},0,0,${act},${dry},0,0,0,0,0,A`;
const CARRY = ["# Farm: TEST DAIRY", HEAD,
  row("a", "Bovaer Premix (BOVAR)", "GROUND CORN (GRDCRN)", "2026-08-20 08:00:00", 1000, 900),
  row("a", "Bovaer Premix (BOVAR)", "Bovaer (Bovaer)", "2026-08-20 08:01:00", 100, 100),
  row("p1", "Premix MILK COWS (PREMIX)", "CANOLA (CANOLA)", "2026-08-21 08:00:00", 1100, 1000),
  row("p1", "Premix MILK COWS (PREMIX)", "Bovaer Premix (BOVAR)", "2026-08-21 08:01:00", 1100, 1000),
  row("h0", "HIGH COW (HIGCOW)", "Premix MILK COWS (PREMIX)", "2026-08-25 08:00:00", 1100, 1000),
  row("h1", "HIGH COW (HIGCOW)", "Premix MILK COWS (PREMIX)", "2026-09-05 08:00:00", 1100, 1000),
  row("h1", "HIGH COW (HIGCOW)", "Corn silage 2025 (CORSL)", "2026-09-05 08:01:00", 3000, 1000),
  row("b", "Bovaer Premix (BOVAR)", "GROUND CORN (GRDCRN)", "2026-09-10 08:00:00", 900, 800),
  row("b", "Bovaer Premix (BOVAR)", "Bovaer (Bovaer)", "2026-09-10 08:01:00", 200, 200),
  row("p2", "Premix MILK COWS (PREMIX)", "CANOLA (CANOLA)", "2026-09-11 08:00:00", 1100, 1000),
  row("p2", "Premix MILK COWS (PREMIX)", "Bovaer Premix (BOVAR)", "2026-09-11 08:01:00", 1100, 1000),
  row("h2", "HIGH COW (HIGCOW)", "Premix MILK COWS (PREMIX)", "2026-09-15 08:00:00", 1100, 1000),
  row("h2", "HIGH COW (HIGCOW)", "Corn silage 2025 (CORSL)", "2026-09-15 08:01:00", 3000, 1000),
].join("\r\n");

test("Load report: BOVAR batch carried forward, premix fed first-in first-out, only the last month counted", () => {
  const { data, info } = C.processReport("Load_report.csv", CARRY, XLSX);
  assert.equal(data.dateRange, "From 9/1/2026 until 9/15/2026");
  close(data.herdDM, 2, "herd DM t (Aug 25 load is history only)");
  // Sep 5 eats the rest of P1 (made with BOVAR A: 500 lb x 10% = 50 lb Bovaer); Sep 15 eats P2 (BOVAR B: 500 x 20% = 100)
  close(data.dose.bovaerAsFed, 0.075, "Bovaer as fed t");
  close(data.dose.mgPerKg, 0.075 * 0.1 / 2 * 1e6, "dose");
  const t = byCode(data);
  close(t.Bovaer, 0.075, "Bovaer DM"); close(t.GRDCRN, 0.425, "GRDCRN"); close(t.CANOLA, 0.5, "CANOLA");
  close(data.afterTotal, data.herdDM, "after = herd");
  assert.match(info.notes.join(" "), /Counted 9\/1\/2026/);
  const s = data.calculator.sources[0];
  close(s.quantityLb, 2200, "PREMIX as fed");
  close(s.inclusionPct, 150 / 1100 * 100, "Bovaer % of BOVAR fed");
  close(data.calculator.mgPerKg, data.dose.mgPerKg, "calculator dose");
});

test("Load report: premix mixed before the month's first BOVAR batch borrows it, with a note", () => {
  const early = LOAD_REPORT.replace(/2026-09-01 08:0/g, "2026-09-02 10:0");
  const { data, info } = C.processReport("Load_report.csv", early, XLSX);
  close(data.dose.mgPerKg, 1000, "dose");
  assert.match(info.notes.join(" "), /mixed before the first Bovaer premix batch/);
});

test("Load report: no BOVAR batches -> BOVAR flagged as missing, no dose", () => {
  const noBovar = LOAD_REPORT.split("\r\n").filter(l => !l.startsWith("b1,")).join("\r\n");
  const { data } = C.processReport("Load_report.csv", noBovar, XLSX);
  assert.deepEqual(data.missing.map(m => m.code), ["BOVAR"]);
  assert.equal(byCode(data).Bovaer, undefined);
  assert.equal(data.dose, null);
});

test("Load report: missing premix batches are flagged, not silently dropped", () => {
  const noPremix = LOAD_REPORT.split("\r\n").filter(l => !l.startsWith("p1,")).join("\r\n");
  const { data } = C.processReport("Load_report.csv", noPremix, XLSX);
  assert.equal(data.missing.length, 1);
  assert.equal(data.missing[0].code, "PREMIX");
  assert.equal(data.dose, null);
});

test("Ingredient/Premix Usage summary is rejected with a clear message", () => {
  const summary = "# Farm: X\nIngredient,Call weight,Loaded weight,Unit,Call weight,Loaded weight,Weight,%,Planned,Actual,Error\nCANOLA (CANOLA),1,1,Ton,1,1,0,0,0,0,0\n";
  assert.throws(() => C.processReport("Ingredient_Premix_Usage.csv", summary, XLSX),
    e => e instanceof C.UserError && /Load report/.test(e.message));
});

test("Legacy derived CSV: BOVAR broken out from its own batch section", () => {
  const csv = [
    "Pen,AvgHeads,RowType,DisplayName,Description,ActualWt_Ton,DryWt_Ton,Units,DateRange",
    "1,100,Ingredient,PREMIX,Premix,1.1,1.0,Ton,From 4/1/2026 until 4/30/2026",
    "1,100,Ingredient,CORSL,Corn silage,9.0,3.0,Ton,From 4/1/2026 until 4/30/2026",
    "1,100,SubTotal,,,10.1,4.0,Ton,From 4/1/2026 until 4/30/2026",
    "PREMIX,,Ingredient,CANOLA,Canola,1.7,1.5,Ton,From 4/1/2026 until 4/30/2026",
    "PREMIX,,Ingredient,BOVAR,Bovaer Premix,0.6,0.5,Ton,From 4/1/2026 until 4/30/2026",
    "BOVAR,,Ingredient,GRDCRN,Ground corn,1.1,0.9,Ton,From 4/1/2026 until 4/30/2026",
    "BOVAR,,Ingredient,Bovaer,Bovaer,0.1,0.1,Ton,From 4/1/2026 until 4/30/2026",
  ].join("\n");
  const { data, info } = C.processReport("April.csv", csv, XLSX);
  assert.equal(info.source, "TMR Tracker ingredient report");
  assert.equal(data.days, 30);
  const t = byCode(data);
  close(t.CANOLA, 0.75, "CANOLA");
  close(t.Bovaer, 0.025, "Bovaer");
  close(t.GRDCRN, 0.225, "GRDCRN");
  close(data.afterTotal, 4, "after total");
  close(data.dose.mgPerKg, 0.025 * 0.1 / 4 * 1e6, "dose");
});

test("Legacy report without a BOVAR section: BOVAR flagged, never guessed", () => {
  const csv = [
    "Pen,AvgHeads,RowType,DisplayName,Description,ActualWt_Ton,DryWt_Ton,Units",
    "1,100,Ingredient,BOVAR,Bovaer Premix,0.6,0.5,Ton",
    "1,100,Ingredient,CORSL,Corn silage,9.0,3.0,Ton",
    "1,100,SubTotal,,,9.6,3.5,Ton",
  ].join("\n");
  const { data } = C.processReport("April.csv", csv, XLSX);
  assert.deepEqual(data.missing.map(m => m.code), ["BOVAR"]);
  assert.equal(data.dose, null);
});

test("Unsupported file types are rejected", () => {
  assert.throws(() => C.processReport("notes.pdf", "", XLSX), C.UserError);
});

test("outputName strips the extension regardless of case", () => {
  assert.equal(C.outputName("Report.XLSX"), "UpLook_Ingredient_Upload_Report");
  assert.equal(C.outputName("Report.Xlsx"), "UpLook_Ingredient_Upload_Report");
  assert.equal(C.outputName("Load_report.csv"), "UpLook_Ingredient_Upload_Load_report");
});

test("Load report: recipe (intended) dose from the batch call weights", () => {
  const withCalls = CARRY.replace(/,0,0,(\d+),(\d+),0,0,0,0,0,A/g, ",$1,$2,$1,$2,0,0,0,0,0,A");
  const { data } = C.processReport("Load_report.csv", withCalls, XLSX);
  // calls = loads here, so intended = actual: 50 lb (BOVAR A) + 100 lb (BOVAR B)
  close(data.dose.intended.bovaerAsFed, 0.075, "intended Bovaer as fed t");
  close(data.dose.intended.mgPerKg, data.dose.mgPerKg, "intended dose");
  assert.equal(C.processReport("Load_report.csv", CARRY, XLSX).data.dose.intended, null);
});

// Premixes report as pdf.js text lines. BOVAR = 10% Bovaer as fed; PREMIX = 20% BOVAR as fed.
const PREMIXES = ["TEST DAIRY", "Premixes", "Generated: 09/17/2026 08:00 AM",
  "Bovaer Premix", "Recipe Display name Density Dry matter % Actual weight Dry weight Price% Not in use%",
  "Bovaer Premix BOVAR 40.00 90.1% 10.00 9.01 1.00",
  "Ingredient Display name Dry matter % Weight/Animal Load % Weight/Animal Load %",
  "1 GROUND CORN GRDCRN 89.0% 9.00 90.0% 8.01 88.9%",
  "2 Bovaer Bovaer 100.0% 1.00 10.0% 1.00 11.1%",
  "Premix MILK COWS", "Recipe Display name Density Dry matter % Actual weight Dry weight Price% Not in use%",
  "Premix MILK COWS PREMIX 40.00 89.2% 10.00 8.92 1.00",
  "Ingredient Display name Dry matter % Weight/Animal Load % Weight/Animal Load %",
  "1 CANOLA CANOLA 89.0% 8.00 80.0% 7.12 79.8%",
  "2 Bovaer Premix BOVAR 90.1% 2.00 20.0% 1.80 20.2%"];
const USAGE = ["# Farm: TEST DAIRY", "Group,# Animals,Ingredient,Actual weight,Dry weight,Unit,Act w/Animal,Dry w/Animal,Unit,Price,Price/Animal",
  "Premix MILK COWS,1 - High ,100,1.00,0.89,Ton,0,0,Lbs,0,0",
  "Corn silage,1 - High ,100,3.00,1.00,Ton,0,0,Lbs,0,0",
  "Hay,9 - Dry ,50,1.00,0.90,Ton,0,0,Lbs,0,0"].join("\n");
const status = r => Object.fromEntries(r.files.map(f => [f.kind, f.status]));

test("Premixes report: recipes, Bovaer premix and the premixes that carry it", () => {
  const p = C.parsePremixes(PREMIXES);
  assert.equal(p.kind, "premixes"); assert.equal(p.generated, "09/17/2026");
  assert.equal(p.bovarCode, "BOVAR"); assert.deepEqual(p.carriers, ["PREMIX"]);
  close(p.recipes.BOVAR.afShare.Bovaer, 0.1, "Bovaer % of BOVAR");
  close(p.recipes.PREMIX.afShare.BOVAR, 0.2, "BOVAR % of PREMIX");
});

test("Ingredient usage + Premixes: intended dose over the pens fed the Bovaer-carrying premix", () => {
  const r = C.analyze([{ name:"Ingredient_usage.csv", content:USAGE }, { name:"Premixes.pdf", content:PREMIXES }], XLSX);
  assert.equal(r.basis, "intended");
  assert.deepEqual(r.result.info.excluded, ["9 - Dry"]);
  close(r.result.data.herdDM, 1.89, "herd DM t");
  close(r.result.data.dose.bovaerAsFed, 0.02, "Bovaer as fed t (1 t x 20% x 10%)");
  close(r.result.data.dose.mgPerKg, 0.02 * 0.1 / 1.89 * 1e6, "dose");
  assert.deepEqual(status(r), { usage:"used", premixes:"used" });
  assert.match(r.guidance.join(" "), /Load report/);
});

test("Each partial file says what else is needed", () => {
  const u = C.analyze([{ name:"Ingredient_usage.csv", content:USAGE }], XLSX);
  assert.equal(u.result, null); assert.equal(status(u).usage, "needs another file");
  assert.match(u.guidance.join(" "), /Premixes report/);
  const p = C.analyze([{ name:"Premixes.pdf", content:PREMIXES }], XLSX);
  assert.equal(p.result, null); assert.match(p.guidance.join(" "), /Ingredient usage/);
  const none = C.analyze([{ name:"notes.txt", content:"" }], XLSX);
  assert.equal(none.result, null); assert.equal(status(none).unknown, "unusable");
  assert.match(none.guidance.join(" "), /Load report/);
});

test("Load report wins; usage + premixes become a cross-check, the summary is flagged unusable", () => {
  const summary = "# Farm: X\nIngredient,Call weight,Loaded weight,Unit\nCANOLA (CANOLA),1,1,Ton\n";
  const r = C.analyze([{ name:"Load_report.csv", content:CARRY }, { name:"Ingredient_usage.csv", content:USAGE },
                       { name:"Premixes.pdf", content:PREMIXES }, { name:"Ingredient_Premix_Usage.csv", content:summary }], XLSX);
  assert.equal(r.basis, "actual");
  close(r.result.data.dose.mgPerKg, C.processReport("Load_report.csv", CARRY, XLSX).data.dose.mgPerKg, "dose from the Load report");
  close(r.result.info.crossCheck.mgPerKg, 0.02 * 0.1 / 1.89 * 1e6, "cross-check");
  assert.deepEqual(status(r), { load:"used", usage:"used", premixes:"used", summary:"unusable" });
});
