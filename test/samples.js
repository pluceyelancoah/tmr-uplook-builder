// Shared by samples.test.js and snapshot.js: turns a real report in raw/ into a comparable summary.
const fs = require("node:fs");
const path = require("node:path");
const XLSX = require("xlsx");
const C = require("../js/core.js");

const RAW = path.join(__dirname, "..", "raw");
const DEFAULT_FRAC = 0.0776;
const r3 = n => Math.round(n * 1000) / 1000;

const sampleFiles = () => fs.existsSync(RAW)
  ? fs.readdirSync(RAW).filter(f => /\.(csv|xlsx?)$/i.test(f)).sort()
  : [];

function summarise(file){
  const full = path.join(RAW, file);
  const content = /\.csv$/i.test(file) ? fs.readFileSync(full, "utf8") : new Uint8Array(fs.readFileSync(full));
  try {
    const { data, info } = C.processReport(file, content, DEFAULT_FRAC, XLSX);
    return { info: { ...info, bovaerFrac: r3(info.bovaerFrac * 100) },
      dateRange: data.dateRange, days: data.days, herdDM: r3(data.herdDM),
      beforeTotal: r3(data.beforeTotal), afterTotal: r3(data.afterTotal),
      missing: data.missing.map(m => m.code), upload: C.uploadRows(data) };
  } catch (e) {
    if (e instanceof C.UserError) return { rejected: e.message };
    throw e;
  }
}

const snapshotPath = file => path.join(RAW, file + ".expected.json");

module.exports = { sampleFiles, summarise, snapshotPath };
