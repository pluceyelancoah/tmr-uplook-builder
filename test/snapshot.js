// Writes raw/<file>.expected.json for every report in raw/. Review the diff before accepting.
const fs = require("node:fs");
const { sampleFiles, summarise, snapshotPath } = require("./samples.js");

for (const f of sampleFiles()) {
  fs.writeFileSync(snapshotPath(f), JSON.stringify(summarise(f), null, 2) + "\n");
  console.log("snapshot:", f);
}
